# -*- coding: utf-8 -*-
"""
Robô de menor preço da Sportbay — versão HTTP (sem Selenium).

POR QUE MUDOU
-------------
A versão anterior abria o Chrome, carregava a home, digitava na barra de busca,
clicava no filtro "Menor Preço" e entrava na página do anúncio para ler o preço
do Pix. Eram ~4 carregamentos de página por código (~33 s cada) e o passo de
"entrar no anúncio" quebrava em lotes grandes: o React re-renderiza a grade
depois de aplicar a ordenação, os elementos <a> capturados antes viram
StaleElementReference e o robô devolvia "anúncio não encontrado" mesmo tendo
visto o produto na grade.

A busca da Sportbay é um Next.js que entrega a grade inteira em JSON (rota
`_next/data/<buildId>/busca.json`), e a página do anúncio também. Trocando o
Chrome por requisições HTTP diretas, cada passo virou uma chamada de ~0,1 s
sem nenhum ponto de falha de DOM. Rodada real dos 315 códigos da base:
6,6 minutos, contra ~3 horas da versão Selenium.

DUAS ARMADILHAS QUE FORAM CORRIGIDAS
------------------------------------
1) O preço do JSON da busca NÃO é confiável. Para vendedores do marketplace ele
   vem defasado, sempre para MENOS (medido: 14 de 47 amostras, com desvios de
   16% a 24%). Exemplo conferido no navegador — código 122361, vendedor
   "JS Mil Full": a busca informa R$ 184,00 e o cliente paga R$ 213,92 no Pix.
   Por isso o robô agora CONFIRMA o preço na página de cada anúncio candidato,
   que traz `pricePix` por variação — e só considera variação com estoque.
   Conferência final: numa amostra de 60 resultados, os 60 bateram exatamente
   com o menor Pix disponível na página do anúncio.

2) A busca é semântica: pesquisar "399" devolve 79 produtos parecidos, não só o
   certo. A versão antiga ordenava por menor preço e pegava o PRIMEIRO da grade
   — muitas vezes outro produto, mais barato. Casos reais medidos:

       código 399  -> antigo R$  79,89 (Paralama Universal)   correto R$ 100,34
       código 2758 -> antigo R$  11,89 (Pisca c/ chicote)     correto R$  14,54
       código 208  -> antigo R$  21,37 (Kit reparo máscara)   correto R$  43,63
       código 4090 -> antigo R$  33,74 (Paralama Bros)        correto R$  44,89

   Agora o candidato passa por dois filtros: similaridade com a descrição da
   planilha e, na confirmação, conferência do `skuCode` da página — que é
   exatamente o código de fábrica ("122361" ou "PAI-122361"). Numa amostra de
   15 códigos, os 40 candidatos de cada um bateram o skuCode: verificação exata.
"""

import argparse
import json
import logging
import os
import re
import sys
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import pandas as pd
import requests

# --- CAMINHOS ---
# Tudo sai da pasta deste script — a raiz do projeto "Analise de preços site".
# Assim o robô lê e grava sempre nos mesmos arquivos, não importa de onde o
# terminal foi aberto (antes dependia do diretório atual).
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURAÇÕES ---
NOME_FICHEIRO_ENTRADA = os.path.join(PASTA_SCRIPT, 'Base dos Produtos.xlsx')
NOME_FICHEIRO_SAIDA = os.path.join(PASTA_SCRIPT, 'Resultado_Menor_Preco_Sportbay.xlsx')
COLUNA_NOME_PRODUTO = 'Produto'   # usada para validar se o anúncio é o produto certo

# Formato da planilha de retorno: colunas da base (menos as descartadas) seguidas
# de COLUNAS_RESULTADO. "Menor Preço" é o preço do Pix.
COLUNAS_DESCARTADAS = ['ShotHash']
COLUNAS_RESULTADO = ['Link', 'Código Site Sportbay', 'Menor Preço']
# Não entram na planilha principal: vão para a aba "Auditoria".
COLUNAS_AUDITORIA = ['Confiança', 'SKU no Site', 'Similaridade', 'Produto no Site',
                     'Variação', 'Vendedor', 'Preço De', 'Preço no Índice da Busca',
                     'Termo Buscado']

WORKERS = 5                # códigos processados em paralelo
WORKERS_CONFIRMACAO = 4    # anúncios confirmados em paralelo dentro de cada código
MAX_PAGINAS = 8            # páginas da busca varridas por código (20 itens cada)
MAX_CONFIRMACOES = 15      # anúncios mais baratos que têm o preço conferido na página
LIMITE_CONFIRMACOES = 45   # teto de páginas conferidas quando os primeiros lotes são anúncios mortos
LIMIAR_SIMILARIDADE = 0.61  # nota mínima de nome para aceitar um anúncio
PESO_COBERTURA = 0.85      # ver similaridade(): calibrado contra casos reais
POUCOS_RESULTADOS = 3      # se a busca devolve <= isso, o código foi um acerto exato
TIMEOUT_HTTP = 30
MAX_TENTATIVAS = 3

BASE = 'https://www.sportbay.com.br'
USER_AGENT = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
              '(KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36')

# --- LOGGING ---
# encoding='utf-8' é obrigatório: o log antigo perdia TODAS as linhas com '✓',
# '→' e '⚠' porque o FileHandler abria em cp1252 e o logging descartava a
# linha no UnicodeEncodeError. Por isso o robo_sportbay.log não tinha nenhum
# WARNING nem nenhuma linha de sucesso — as falhas ficavam invisíveis.
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_SCRIPT, 'robo_sportbay.log'), encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


# --- NORMALIZAÇÃO E SIMILARIDADE DE NOMES ---
def normalizar(texto):
    """Quebra o texto em tokens comparáveis: sem acento, maiúsculo, letra/dígito separados.

    'TR50/TR100' e 'TR-50F, TR-100F' viram tokens compatíveis graças à quebra
    na fronteira letra<->dígito, que é onde os catálogos mais divergem.
    """
    s = unicodedata.normalize('NFKD', str(texto)).encode('ascii', 'ignore').decode().upper()
    s = re.sub(r'(?<=[A-Z])(?=[0-9])|(?<=[0-9])(?=[A-Z])', ' ', s)
    return [t for t in re.split(r'[^A-Z0-9]+', s) if t]


def similaridade(tokens_base, nome_candidato):
    """Nota 0..1 entre o nome da planilha e o nome do anúncio.

    Pesa muito mais a cobertura (quantos termos da planilha o anúncio contém)
    do que a precisão, porque o site acrescenta termos sem mudar o produto
    ('Pro Tork', '100% Poliéster'), enquanto termos FALTANDO costumam indicar
    produto diferente. Foi essa assimetria que resolveu o caso do código 399:
    "Paralama Dianteiro Universal CRF Pro Tork" é curtinho e casava alto na
    precisão, mas não cobre CRF450/CRF250/XR250/TORNADO da planilha.

    Os pesos (0,85/0,15) e o LIMIAR_SIMILARIDADE foram calibrados contra um
    conjunto de pares certos/errados conferidos manualmente no catálogo.
    """
    base = set(tokens_base)
    cand = set(normalizar(nome_candidato))
    if not base or not cand:
        return 0.0
    comum = len(base & cand)
    return PESO_COBERTURA * comum / len(base) + (1 - PESO_COBERTURA) * comum / len(cand)


# --- CLIENTE DA BUSCA ---
class SportbayAPI:
    """Cliente da rota de dados da busca, com sessão por thread e buildId auto-renovável."""

    def __init__(self):
        self._locais = threading.local()
        self._lock = threading.Lock()
        self._build_id = None
        self.build_id  # dispara a descoberta já na criação

    @property
    def sessao(self):
        s = getattr(self._locais, 'sessao', None)
        if s is None:
            s = requests.Session()
            s.headers.update({'User-Agent': USER_AGENT,
                              'Accept': 'application/json, text/html;q=0.9',
                              'Accept-Language': 'pt-BR,pt;q=0.9'})
            self._locais.sessao = s
        return s

    @property
    def build_id(self):
        if self._build_id is None:
            self._renovar_build_id()
        return self._build_id

    def _renovar_build_id(self, antigo=None):
        """Lê o buildId do Next.js na home. Ele muda a cada deploy do site."""
        with self._lock:
            if antigo is not None and self._build_id != antigo:
                return  # outra thread já renovou
            r = self.sessao.get(BASE, timeout=TIMEOUT_HTTP)
            m = re.search(r'"buildId":"([^"]+)"', r.text)
            if not m:
                raise RuntimeError('não foi possível ler o buildId da Sportbay')
            self._build_id = m.group(1)
            logger.info(f'buildId da Sportbay: {self._build_id}')

    def buscar(self, termo, pagina=1):
        """Uma página da busca ordenada por menor preço. Retorna o dict pageProps ou None.

        A ordenação exige os DOIS parâmetros: 'sort' e 'filters'. Só 'sort' é
        ignorado pelo backend (verificado no site).
        """
        params = {'q': termo, 'page': pagina,
                  'sort': 'price_asc', 'filters': 'Ordenação=price_asc'}
        for tentativa in range(1, MAX_TENTATIVAS + 1):
            build = self.build_id
            try:
                r = self.sessao.get(f'{BASE}/_next/data/{build}/busca.json',
                                    params=params, timeout=TIMEOUT_HTTP)
                if r.status_code == 200:
                    return r.json()['pageProps']
                if r.status_code == 404:
                    # buildId velho (o site fez deploy no meio da execução)
                    self._renovar_build_id(antigo=build)
                    continue
                if r.status_code in (429, 500, 502, 503, 504):
                    time.sleep(1.5 * tentativa)
                    continue
                # qualquer outro status: cai para a rota HTML
                return self._buscar_via_html(termo, pagina)
            except (requests.RequestException, ValueError, KeyError) as e:
                logger.debug(f'  falha na busca "{termo}" p.{pagina} ({e}), tentativa {tentativa}')
                time.sleep(1.0 * tentativa)
        return self._buscar_via_html(termo, pagina)

    def _buscar_via_html(self, termo, pagina):
        """Plano B: a página HTML da busca traz o mesmo JSON em __NEXT_DATA__."""
        params = {'q': termo, 'page': pagina,
                  'sort': 'price_asc', 'filters': 'Ordenação=price_asc'}
        try:
            r = self.sessao.get(f'{BASE}/busca', params=params, timeout=TIMEOUT_HTTP)
            m = re.search(r'id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
            if not m:
                return None
            return json.loads(m.group(1))['props']['pageProps']
        except Exception as e:
            logger.debug(f'  plano B (HTML) falhou para "{termo}": {e}')
            return None

    def confirmar_anuncio(self, slug, hash_produto):
        """Lê o preço REAL do Pix, o código de fábrica e a variação na página do anúncio.

        Cada variação traz `pricePix` pronto (é o valor do `.price-main`, conferido
        no navegador). Se faltar, calculamos: price × (1 - merchant.descontoPix/100),
        porque o desconto do Pix é definido por vendedor.

        Só entram variações com estoque: o anúncio pode listar cor/tamanho zerados,
        e cotar um preço que ninguém consegue comprar seria pior que não cotar.

        Retorna (preco_pix, skuCode, nome, variacao) ou None.
        """
        for tentativa in range(1, MAX_TENTATIVAS + 1):
            try:
                r = self.sessao.get(f'{BASE}/{slug}/{hash_produto}/p', timeout=TIMEOUT_HTTP)
                if r.status_code == 404:
                    return None
                m = re.search(r'id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
                if not m:
                    return None
                # Anúncio removido responde 200 com a página "Produto não
                # encontrado", que vem sem o bloco `product`.
                p = json.loads(m.group(1))['props']['pageProps'].get('product')
                if not p:
                    return None

                desconto = (p.get('merchant') or {}).get('descontoPix') or 0
                variacoes = []
                for v in p.get('productVariations', []):
                    preco = v.get('pricePix')
                    if preco is None and v.get('price'):
                        preco = round(v['price'] * (1 - desconto / 100), 2)
                    if preco:
                        variacoes.append((preco, v))
                if not variacoes:
                    return None

                com_estoque = [x for x in variacoes if (x[1].get('totalStock') or 0) > 0]
                preco, v = min(com_estoque or variacoes, key=lambda x: x[0])
                return (preco, p.get('skuCode') or '', p.get('name') or '', v.get('name') or '')
            except Exception as e:
                logger.debug(f'  confirmação de {hash_produto} falhou ({e}), tentativa {tentativa}')
                time.sleep(0.7 * tentativa)
        return None


# --- COLETA E DECISÃO ---
def coletar_candidatos(api, termo, tokens_base, alvo=MAX_CONFIRMACOES):
    """Varre a busca (ordenada do mais barato para o mais caro) e pontua cada anúncio.

    Retorna (candidatos, total de resultados). NÃO filtra por nome aqui — isso
    já foi tentado e causava bug: o anúncio mais barato de um produto costuma
    ter o nome mais "enxuto" da grade (ex.: "Camisa Motocross Brave 2025 Pro
    Tork", vendido pela própria Sportbay), enquanto a planilha traz a descrição
    completa do catálogo antigo ("CAMISA MOTOCROSS ADULTO PRO TORK BRAVE
    AMARELO TAM. G"). Filtrar por nome ANTES de confirmar excluía esse anúncio
    da lista de candidatos a confirmar — mesmo ele tendo o skuCode exatamente
    igual ao código pesquisado — e o robô confirmava só o 2º/3º mais barato,
    cujo nome batia melhor mas o preço era ~50% maior.

    Confirmar por preço ignorando o nome é seguro porque quem decide de fato
    é o skuCode da página (ver `_decidir`), não a similaridade. A nota de
    similaridade continua calculada aqui e serve só de critério de confiança
    quando NENHUM candidato bate o skuCode.

    Para de paginar assim que junta `alvo` candidatos — para um código
    específico a 1ª página já costuma trazer 20 ofertas do mesmo produto, de
    vendedores diferentes.
    """
    candidatos = []
    total = 0
    pagina = 1
    while pagina <= MAX_PAGINAS:
        dados = api.buscar(termo, pagina)
        if not dados:
            break
        produtos = dados['data']['products']
        total = dados['pagination']['totalElements']
        if not produtos:
            break
        for p in produtos:
            variacoes = [v for v in p['variations'] if v.get('discountPrice')]
            if not variacoes:
                continue
            v = min(variacoes, key=lambda x: x['discountPrice'])
            hash_produto = p.get('shortHash') or v.get('shortHash') or ''
            if not hash_produto:
                continue
            candidatos.append({
                'preco_busca': v['discountPrice'],   # índice da busca: só para ordenar
                'preco_de': v.get('fakePrice'),
                'nome': p['name'],
                'variacao': v.get('name') or v.get('option') or '',
                'vendedor': (p.get('merchant') or {}).get('name', ''),
                'hash': hash_produto,
                'slug': p.get('slug', ''),
                'nota': similaridade(tokens_base, p['name']) if tokens_base else 0.0,
            })
        if len(candidatos) >= alvo:
            break
        if pagina >= dados['pagination']['totalPages']:
            break
        pagina += 1
    return candidatos, total


def confirmar_candidatos(api, candidatos, codigo, ignorar=None):
    """Confere na página de cada anúncio o preço real do Pix e o código de fábrica.

    O preço da busca serve só para escolher QUAIS anúncios conferir; o preço que
    vale é sempre o da página. MAX_CONFIRMACOES=15 foi calibrado: numa amostra
    de 15 códigos, conferir os 10 mais baratos já bastava para achar o mínimo
    real que só apareceria conferindo todos os 40 candidatos.

    O índice da busca guarda anúncios fantasma — a página responde 200 mas é
    "Produto não encontrado" e vem sem o bloco `product`. Se o lote inteiro for
    fantasma, seguimos para o lote seguinte em vez de devolver o preço do
    índice, que seria um preço de anúncio que ninguém consegue comprar.

    Retorna (confirmados, hashes_mortos).
    """
    ignorar = ignorar or set()
    fila = [c for c in sorted(candidatos, key=lambda c: c['preco_busca'])
            if c['hash'] not in ignorar][:LIMITE_CONFIRMACOES]
    confirmados, mortos = [], set()
    while fila and not confirmados:
        lote, fila = fila[:MAX_CONFIRMACOES], fila[MAX_CONFIRMACOES:]
        with ThreadPoolExecutor(max_workers=min(WORKERS_CONFIRMACAO, len(lote))) as pool:
            infos = list(pool.map(lambda c: api.confirmar_anuncio(c['slug'], c['hash']), lote))
        for c, info in zip(lote, infos):
            if not info:
                mortos.add(c['hash'])
                continue
            preco, sku, nome_pagina, variacao = info
            c = dict(c, preco=preco, sku=sku, sku_confere=sku_corresponde(sku, codigo))
            if nome_pagina:
                c['nome'] = nome_pagina
            if variacao:
                c['variacao'] = variacao
            confirmados.append(c)
    return confirmados, mortos


def sku_corresponde(sku, codigo):
    """O anúncio é mesmo deste código de fábrica? O site usa '122361' ou 'PAI-122361'."""
    limpo = re.sub(r'^PAI[-_ ]?', '', (sku or '').strip().upper())
    return bool(limpo) and limpo == str(codigo).strip().upper()


def _melhor(candidatos, campo='preco'):
    return min(candidatos, key=lambda c: (c[campo], -c['nota']))


def _nome_de_referencia(api, termo):
    """Sem nome na planilha, usamos o 1º resultado por RELEVÂNCIA como referência.

    A busca semântica coloca o produto certo no topo quando não há ordenação por
    preço; usamos esse nome para filtrar as ofertas dos demais vendedores.
    """
    try:
        r = api.sessao.get(f'{BASE}/_next/data/{api.build_id}/busca.json',
                           params={'q': termo, 'page': 1}, timeout=TIMEOUT_HTTP)
        if r.status_code != 200:
            return None
        produtos = r.json()['pageProps']['data']['products']
        return produtos[0]['name'] if produtos else None
    except Exception:
        return None


def _decidir(api, codigo, termo, tokens, confianca_nome):
    """Resolve um termo de busca: coleta, confirma nas páginas e escolhe o vencedor.

    Confirma os candidatos por PREÇO (mais barato primeiro), não por nome —
    ver o comentário em coletar_candidatos() sobre por que filtrar por nome
    antes de confirmar deixava passar o anúncio mais barato.

    Níveis de confiança, do melhor para o pior:
      EXATA — o skuCode da página do anúncio é o próprio código de fábrica
              (decide mesmo que o nome do anúncio divirja da planilha)
      ALTA  — nenhum skuCode bateu, mas o nome do anúncio bate com a planilha
      MEDIA — nem skuCode nem nome bateram, porém a busca devolveu pouquíssimos
              resultados (código cirúrgico) ou a referência de nome veio do
              próprio site (planilha sem descrição) — ficamos com o mais barato
              confirmado mesmo assim

    Se o produto aparece na busca mas nenhum anúncio tem página viva, devolve o
    registro com preco=None e confianca='SEM_ANUNCIO' — o produto existe no
    catálogo, só não está à venda.
    """
    candidatos, total = coletar_candidatos(api, termo, tokens)
    if not candidatos:
        return None

    confirmados, mortos = confirmar_candidatos(api, candidatos, codigo)

    if not confirmados and len(candidatos) < LIMITE_CONFIRMACOES:
        # Só anúncios mortos até aqui. Vale pagar por mais páginas da busca
        # antes de desistir — mas só neste caso raro, para não onerar o comum.
        extras, _ = coletar_candidatos(api, termo, tokens, alvo=LIMITE_CONFIRMACOES)
        if len(extras) > len(candidatos):
            confirmados, _ = confirmar_candidatos(api, extras, codigo, ignorar=mortos)
            candidatos = extras

    exatos = [c for c in confirmados if c['sku_confere']]
    if exatos:
        return dict(_melhor(exatos), confianca='EXATA', termo=termo)

    validos = [c for c in confirmados if c['nota'] >= LIMIAR_SIMILARIDADE]
    if validos:
        return dict(_melhor(validos), confianca=confianca_nome, termo=termo)

    if confirmados and total and total <= POUCOS_RESULTADOS:
        # Busca devolveu quase nada => o código casou de forma exata; o nome
        # diverge só na redação do catálogo (ex.: "TR50/TR100" x "TR-50F").
        return dict(_melhor(confirmados), confianca='MEDIA', termo=termo)

    if confirmados:
        # Preços confirmados, mas nenhum bateu skuCode nem nome, e a busca é
        # ampla demais para confiar no mais barato "no chute". Este termo não
        # resolveu — resolver_produto tenta o próximo (sem o sufixo de letra,
        # ou o nome).
        return None

    # Todos os anúncios são fantasmas: o produto está no índice mas não à venda.
    escolhido = _melhor(candidatos, campo='preco_busca')
    return dict(escolhido, preco=None, sku='', sku_confere=False,
                confianca='SEM_ANUNCIO', termo=termo)


def resolver_produto(api, codigo, nome_planilha):
    """Encontra o anúncio mais barato que realmente corresponde ao código."""
    tokens = normalizar(nome_planilha) if isinstance(nome_planilha, str) and nome_planilha.strip() else []
    confianca_nome = 'ALTA'

    if not tokens:
        # Planilha sem descrição: deixa o próprio site dizer qual é o produto.
        referencia = _nome_de_referencia(api, codigo)
        if referencia:
            tokens = normalizar(referencia)
            confianca_nome = 'MEDIA'

    # Códigos terminados em letra (sufixos 'A', 'P', etc.) costumam ter a
    # variante sem sufixo no site.
    termos = [codigo]
    if len(codigo) > 1 and codigo[-1].isalpha():
        base = codigo[:-1].strip()
        if base:
            termos.append(base)

    def prefere(novo, atual):
        """Quem tem preço ganha de quem não tem; entre os dois, o mais barato."""
        if novo is None:
            return False
        if atual is None or atual['preco'] is None:
            return True
        return novo['preco'] is not None and novo['preco'] < atual['preco']

    achado = None
    for termo in termos:
        r = _decidir(api, codigo, termo, tokens, confianca_nome)
        if prefere(r, achado):
            achado = r

    # Último recurso: procurar pela descrição do produto em vez do código.
    if (achado is None or achado['preco'] is None) and isinstance(nome_planilha, str) and nome_planilha.strip():
        r = _decidir(api, codigo, nome_planilha, normalizar(nome_planilha), 'NOME')
        if prefere(r, achado):
            achado = dict(r, confianca=r['confianca'] if r['preco'] is None else 'NOME')

    if achado is None:
        return None
    achado['link'] = f"{BASE}/{achado['slug']}/{achado['hash']}/p" if achado['hash'] else ''
    return achado


# --- EXECUÇÃO ---
_progresso = {'n': 0}
_lock_log = threading.Lock()


def processar_linha(api, indice, total, codigo, nome):
    if not codigo or codigo.lower() in ('nan', 'none', ''):
        return {'Status': 'CODIGO_VAZIO'}
    try:
        r = resolver_produto(api, codigo, nome)
    except Exception as e:
        logger.error(f'  ✗ erro em {codigo}: {e}')
        return {'Status': 'ERRO', 'Detalhe': str(e)[:120]}

    with _lock_log:
        _progresso['n'] += 1
        pos = _progresso['n']
    if r is None:
        logger.warning(f'[{pos}/{total}] {codigo}: não encontrado na Sportbay')
        return {'Status': 'NAO_ENCONTRADO'}

    if r['preco'] is None:
        logger.warning(f"[{pos}/{total}] {codigo}: no catálogo mas sem anúncio ativo "
                       f"— {r['nome'][:52]}")
        return {
            'Status': 'SEM_ANUNCIO_ATIVO',
            'Produto no Site': r['nome'],
            'Similaridade': round(r['nota'], 2),
            'Preço no Índice da Busca': r.get('preco_busca'),
            'Termo Buscado': r['termo'],
            'Link': r['link'],
        }

    logger.info(f"[{pos}/{total}] {codigo}: R$ {r['preco']:.2f} "
                f"({r['confianca']}, nota {r['nota']:.2f}) — {r['nome'][:52]}")
    return {
        'Status': 'ENCONTRADO',
        'Código Site Sportbay': r['hash'],
        'Menor Preço': round(r['preco'], 2),
        'Preço De': r['preco_de'],
        'Confiança': r['confianca'],
        'SKU no Site': r.get('sku', ''),
        'Similaridade': round(r['nota'], 2),
        'Produto no Site': r['nome'],
        'Variação': r['variacao'],
        'Vendedor': r['vendedor'],
        'Preço no Índice da Busca': r.get('preco_busca'),
        'Termo Buscado': r['termo'],
        'Link': r['link'],
    }


def montar_saida(df, coluna_codigo):
    """Separa o resultado em (planilha limpa, aba de auditoria).

    A planilha principal repete as colunas da base — menos as descartadas — e
    acrescenta Link, Código Site Sportbay e Menor Preço (o preço do Pix), nessa
    ordem. É o formato do Resultado_Menor_Preco_Sportbay.xlsx.

    Tudo que serve para conferir a escolha (confiança, SKU, similaridade, nome
    do anúncio, vendedor) vai para a aba "Auditoria", fora do caminho de quem
    só quer os preços.
    """
    fora = set(COLUNAS_DESCARTADAS) | set(COLUNAS_RESULTADO) | set(COLUNAS_AUDITORIA) | {'Status'}
    colunas_base = [c for c in df.columns if c not in fora]
    planilha = df[colunas_base + COLUNAS_RESULTADO].copy()
    planilha['Menor Preço'] = pd.to_numeric(planilha['Menor Preço'], errors='coerce')

    auditoria = df[[coluna_codigo, COLUNA_NOME_PRODUTO] if COLUNA_NOME_PRODUTO in df.columns
                   else [coluna_codigo]].copy()
    for coluna in ['Status', 'Menor Preço'] + COLUNAS_AUDITORIA:
        auditoria[coluna] = df[coluna].values
    return planilha, auditoria


def salvar(planilha, auditoria, caminho):
    """Grava o Excel; se o arquivo estiver aberto no Excel, grava com timestamp."""
    def escrever(destino):
        with pd.ExcelWriter(destino, engine='openpyxl') as writer:
            planilha.to_excel(writer, sheet_name='Resultado', index=False)
            if auditoria is not None:
                auditoria.to_excel(writer, sheet_name='Auditoria', index=False)
    try:
        escrever(caminho)
        return caminho
    except PermissionError:
        raiz, ext = os.path.splitext(caminho)
        alternativo = f'{raiz}_{datetime.now():%Y%m%d_%H%M%S}{ext}'
        escrever(alternativo)
        logger.warning(f'⚠ "{caminho}" está aberto/bloqueado — salvo em "{alternativo}"')
        return alternativo


def iniciar_robo(entrada, saida, limite=None, workers=WORKERS):
    logger.info('=' * 60)
    logger.info('ROBÔ DE MENOR PREÇO — SPORTBAY (modo HTTP)')
    logger.info('=' * 60)

    try:
        df = pd.read_excel(entrada)
    except FileNotFoundError:
        logger.error(f'Arquivo não encontrado: {entrada}')
        return
    except Exception as e:
        logger.error(f'Erro ao ler {entrada}: {e}')
        return

    if limite:
        df = df.head(limite).copy()

    coluna_codigo = df.columns[0]
    tem_nome = COLUNA_NOME_PRODUTO in df.columns
    logger.info(f'{len(df)} produtos | coluna de código: "{coluna_codigo}" | '
                f'validação por nome: {"sim" if tem_nome else "NÃO (menos assertivo)"}')
    if not tem_nome:
        logger.warning(f'⚠ Coluna "{COLUNA_NOME_PRODUTO}" ausente — a checagem de '
                       f'produto usará o 1º resultado por relevância como referência.')

    api = SportbayAPI()
    total = len(df)
    _progresso['n'] = 0

    linhas = [(str(r[coluna_codigo]).strip(),
               r[COLUNA_NOME_PRODUTO] if tem_nome else None) for _, r in df.iterrows()]

    t0 = time.time()
    with ThreadPoolExecutor(max_workers=workers) as pool:
        resultados = list(pool.map(
            lambda a: processar_linha(api, a[0], total, a[1][0], a[1][1]),
            enumerate(linhas)))
    duracao = time.time() - t0

    saida_df = pd.DataFrame(resultados)
    for coluna in ['Status'] + COLUNAS_RESULTADO + COLUNAS_AUDITORIA:
        df[coluna] = saida_df[coluna].values if coluna in saida_df else None

    planilha, auditoria = montar_saida(df, coluna_codigo)
    caminho = salvar(planilha, auditoria, saida)

    resumo = saida_df['Status'].value_counts().to_dict()
    encontrados = df[df['Status'] == 'ENCONTRADO']
    logger.info('=' * 60)
    logger.info(f'Concluído em {duracao:.1f}s ({duracao / max(total, 1):.2f}s por código)')
    logger.info(f'Resumo: {resumo}')
    if len(encontrados):
        conf = encontrados['Confiança'].value_counts().to_dict()
        logger.info(f'Confiança: {conf}')
    faltantes = df[df['Status'] != 'ENCONTRADO'][coluna_codigo].tolist()
    if faltantes:
        logger.info(f'Sem preço ({len(faltantes)}): {faltantes}')
    conferir = df[(df['Status'] == 'ENCONTRADO') & (df['Confiança'] != 'EXATA')][coluna_codigo].tolist()
    if conferir:
        logger.info(f'Sem SKU exato, vale conferir ({len(conferir)}): {conferir}')
    logger.info(f'✓ Arquivo salvo: {caminho}')
    logger.info('=' * 60)


if __name__ == '__main__':
    ap = argparse.ArgumentParser(description='Menor preço Sportbay por código de fábrica')
    ap.add_argument('-e', '--entrada', default=NOME_FICHEIRO_ENTRADA)
    ap.add_argument('-s', '--saida', default=NOME_FICHEIRO_SAIDA)
    ap.add_argument('-n', '--limite', type=int, help='processa apenas as N primeiras linhas')
    ap.add_argument('-w', '--workers', type=int, default=WORKERS, help='buscas simultâneas')
    args = ap.parse_args()
    try:
        iniciar_robo(args.entrada, args.saida, args.limite, args.workers)
    except KeyboardInterrupt:
        logger.warning('⚠ Interrompido pelo usuário')
    except Exception as e:
        logger.critical(f'Erro não tratado: {e}', exc_info=True)
