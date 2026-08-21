# -*- coding: utf-8 -*-
"""
Robô de ordenação dos produtos nos carrosséis (coleções) do painel Sportbay.

O QUE ELE FAZ
-------------
Lê a planilha "ORDENAÇÃO DE PRODUTOS.xlsx" (coluna A = SKU do Pai,
coluna B = Ordenação, coluna C = Coleção), entra no painel do marketplace e,
para cada linha, abre o produto correspondente dentro da coleção e grava o
número da coluna B no campo "Ordenação", salvando em "Atualizar".

CAMINHO PERCORRIDO
------------------
    login  ->  /dashboard/categories  (menu CONFIGURAÇÃO > Categoria > Lista)
           ->  seção "Coleções", botão com o nome exato da coluna C
           ->  tabela "Produtos da categoria", linha cujo SKU casa com a coluna A
           ->  página do produto, campo name="ordering"  ->  botão "Atualizar"

DUAS COISAS QUE PRECISAM DE CUIDADO
-----------------------------------
1) O SKU da planilha e o SKU do painel divergem. No painel aparecem formas como
   "PAI-119603A", "119610" e "4902"; na planilha, "119603A", "119610A" e "4902A".
   Por isso os dois lados passam pela mesma normalização (tira o prefixo "PAI-"
   e o sufixo "A") antes da comparação. Nos 16 produtos da coleção
   "Oferta do Dia" essa regra casa 16 de 16.

2) O link do nome do produto tem target="_blank": clicar nele abre uma aba nova.
   O robô lê o href e navega na mesma aba — mesmo destino, sem gerenciar abas.

O campo "Ordenação" é um input controlado pelo React; escrever com send_keys
gera os eventos de teclado que o React escuta, então o valor não é revertido no
re-render (mudar .value por JavaScript seria descartado).

USO
---
    python robo_ordenacao.py            # roda a planilha inteira e salva
    python robo_ordenacao.py --teste    # só a 1ª linha, SEM salvar (ensaio)
    python robo_ordenacao.py --linhas 5 # só as 5 primeiras linhas
"""

import argparse
import logging
import os
import re
import sys
import time
import unicodedata
from collections import OrderedDict

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# --- CONFIGURAÇÕES ---
PASTA = os.path.dirname(os.path.abspath(__file__))
NOME_PLANILHA = 'ORDENAÇÃO DE PRODUTOS.xlsx'

COLUNA_SKU = 'SKU do Pai'
COLUNA_ORDEM = 'Ordenação'
COLUNA_COLECAO = 'Coleção'

BASE = 'https://www.marketplace.sportbay.com.br'
URL_LOGIN = f'{BASE}/auth/jwt/login?returnTo=%2Fdashboard%2Fauction%2Flist'
URL_CATEGORIAS = f'{BASE}/dashboard/categories'

TEMPO_ESPERA = 40          # timeout padrão dos WebDriverWait, em segundos
PAUSA_APOS_SALVAR = 2.5    # respiro depois do "Atualizar" antes da próxima linha

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)-7s | %(message)s',
    datefmt='%H:%M:%S',
    handlers=[
        logging.FileHandler(os.path.join(PASTA, 'robo_ordenacao.log'), encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger('ordenacao')


# ---------------------------------------------------------------- utilidades

def normalizar_sku(valor):
    """Reduz o SKU à forma comparável: sem prefixo PAI- e sem sufixo A.

    "PAI-119603A", "119603A" e "119603" viram todos "119603".
    """
    if valor is None:
        return ''
    texto = str(valor).strip().upper()
    texto = re.sub(r'\.0$', '', texto)          # 125622.0 -> 125622 (Excel numérico)
    texto = re.sub(r'^PAI[-_ ]*', '', texto)
    texto = re.sub(r'A$', '', texto)
    return texto.strip()


def normalizar_nome(valor):
    """Nome comparável: sem acento, sem espaço duplo, minúsculo.

    Só é usado como rede de segurança quando o nome exato da coluna C não
    aparece na lista — o casamento preferido continua sendo o exato.
    """
    texto = unicodedata.normalize('NFKD', str(valor or ''))
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r'\s+', ' ', texto).strip().lower()


def ler_planilha(caminho):
    df = pd.read_excel(caminho)
    faltando = [c for c in (COLUNA_SKU, COLUNA_ORDEM, COLUNA_COLECAO) if c not in df.columns]
    if faltando:
        raise SystemExit(f'Colunas ausentes na planilha: {faltando}. Encontradas: {list(df.columns)}')

    linhas = []
    for i, row in df.iterrows():
        sku = row[COLUNA_SKU]
        ordem = row[COLUNA_ORDEM]
        colecao = row[COLUNA_COLECAO]
        if pd.isna(sku) or pd.isna(ordem) or pd.isna(colecao):
            log.warning('Linha %d ignorada (campo vazio): %s | %s | %s', i + 2, sku, ordem, colecao)
            continue
        linhas.append({
            'linha': i + 2,                       # numeração como aparece no Excel
            'sku_planilha': str(sku).strip(),
            'sku': normalizar_sku(sku),
            'ordem': str(int(ordem)) if float(ordem).is_integer() else str(ordem).strip(),
            'colecao': str(colecao).strip(),
        })
    return linhas


# ------------------------------------------------------------------ navegador

def abrir_navegador():
    """Chrome sem janela: o robô roda só no terminal.

    Sem janela não existe "maximizar", e o tamanho padrão do headless é pequeno
    demais — a captura de tela do modo teste sairia cortada e elementos fora da
    viewport atrapalhariam os cliques. Daí o --window-size explícito.
    """
    opcoes = Options()
    opcoes.add_argument('--headless=new')
    opcoes.add_argument('--window-size=1920,1080')
    opcoes.add_argument('--disable-gpu')
    opcoes.add_argument('--no-sandbox')
    opcoes.add_argument('--disable-dev-shm-usage')
    opcoes.add_argument('--disable-blink-features=AutomationControlled')
    opcoes.add_argument('--log-level=3')
    opcoes.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(options=opcoes)   # Selenium Manager resolve o chromedriver
    driver.set_page_load_timeout(90)
    return driver


def fazer_login(driver, usuario, senha):
    log.info('Abrindo a tela de login...')
    driver.get(URL_LOGIN)
    espera = WebDriverWait(driver, TEMPO_ESPERA)

    campo_usuario = espera.until(EC.element_to_be_clickable((By.NAME, 'email')))
    campo_usuario.clear()
    campo_usuario.send_keys(usuario)

    campo_senha = driver.find_element(By.NAME, 'password')
    campo_senha.clear()
    campo_senha.send_keys(senha)

    driver.find_element(By.CSS_SELECTOR, 'button[type="submit"]').click()

    # O painel só considera logado quando sai de /auth/ e o menu lateral aparece.
    espera.until(lambda d: '/auth/' not in d.current_url)
    espera.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href="/dashboard"]')))
    log.info('Login concluído (%s).', driver.current_url)


# ---------------------------------------------------------------- coleções

def abrir_lista_de_categorias(driver):
    """Equivale ao caminho CONFIGURAÇÃO > Categoria > Lista do menu lateral.

    O título "Coleções" aparece antes da lista: por alguns segundos o painel
    mostra um esqueleto vazio no lugar dos botões. Esperar só pelo título faz o
    robô procurar a coleção numa lista que ainda não existe — a espera é pelos
    botões já renderizados.
    """
    driver.get(URL_CATEGORIAS)
    WebDriverWait(driver, TEMPO_ESPERA).until(
        EC.presence_of_element_located((By.XPATH, "//h6[normalize-space()='Coleções']"))
    )
    WebDriverWait(driver, TEMPO_ESPERA).until(
        lambda d: len(d.find_elements(
            By.XPATH, "//h6[normalize-space()='Coleções']/following-sibling::div//button")) > 0
    )


def abrir_colecao(driver, nome):
    """Clica no botão da coleção dentro da seção 'Coleções' e espera a edição abrir."""
    abrir_lista_de_categorias(driver)

    xpath_exato = (
        "//h6[normalize-space()='Coleções']/following-sibling::div"
        f"//button[normalize-space()={xpath_literal(nome)}]"
    )
    botoes = driver.find_elements(By.XPATH, xpath_exato)

    if not botoes:
        # Rede de segurança: mesma coleção com acento/caixa/espaço diferentes.
        alvo = normalizar_nome(nome)
        todos = driver.find_elements(
            By.XPATH, "//h6[normalize-space()='Coleções']/following-sibling::div//button"
        )
        botoes = [b for b in todos if normalizar_nome(b.text) == alvo]
        if botoes:
            log.warning('Coleção "%s" casou por aproximação com "%s".', nome, botoes[0].text.strip())

    if not botoes:
        raise LookupError(f'Coleção "{nome}" não encontrada na seção Coleções.')

    driver.execute_script('arguments[0].scrollIntoView({block:"center"});', botoes[0])
    botoes[0].click()

    WebDriverWait(driver, TEMPO_ESPERA).until(
        EC.presence_of_element_located((By.XPATH, "//h6[normalize-space()='Produtos da categoria']"))
    )
    log.info('Coleção "%s" aberta (%s).', nome, driver.current_url)


def xpath_literal(texto):
    """Texto seguro para XPath, inclusive com aspas simples no nome."""
    if "'" not in texto:
        return f"'{texto}'"
    partes = texto.split("'")
    return "concat(" + ", \"'\", ".join(f"'{p}'" for p in partes) + ")"


JS_SELETOR_POR_PAGINA = """
const h = [...document.querySelectorAll('h6')]
    .find(e => e.textContent.trim() === 'Produtos da categoria');
if (!h) return null;
const pag = [...document.querySelectorAll('.MuiTablePagination-root')].find(
    p => h.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_FOLLOWING);
if (!pag) return null;
const rotulo = pag.querySelector('.MuiTablePagination-displayedRows');
return {
    seletor: pag.querySelector('[role="button"][aria-haspopup="listbox"]'),
    porPagina: parseInt((pag.querySelector('input.MuiSelect-nativeInput') || {}).value || '0', 10),
    total: parseInt(((rotulo ? rotulo.textContent : '').match(/de\\s+(\\d+)/) || [0, 0])[1], 10)
};
"""


def ampliar_pagina_de_produtos(driver):
    """Sobe o "Número de produtos por página" da tabela de produtos associados.

    A tabela abre com 10 linhas por página. Numa coleção com 16 produtos, os 6
    últimos SKUs da planilha simplesmente não existiriam para o robô — daria
    "não encontrado" num produto que está lá. Subindo para a maior opção (1000),
    a coleção inteira cabe numa página só e a leitura fica sendo uma só varredura.
    """
    dados = driver.execute_script(f'return (function() {{{JS_SELETOR_POR_PAGINA}}})();')
    if not dados or not dados.get('seletor'):
        return
    if dados['total'] <= dados['porPagina']:
        return

    dados['seletor'].click()
    opcoes = WebDriverWait(driver, TEMPO_ESPERA).until(
        lambda d: d.find_elements(By.CSS_SELECTOR, 'ul[role="listbox"] li') or False
    )
    opcoes[-1].click()      # a maior opção da lista (5, 10, 25, 50, 100, 500, 1000)

    WebDriverWait(driver, TEMPO_ESPERA).until(
        lambda d: (d.execute_script(f'return (function() {{{JS_SELETOR_POR_PAGINA}}})();') or {})
        .get('porPagina', 0) > dados['porPagina']
    )
    log.info('Produtos por página: %d -> maior opção (%d produtos na coleção).',
             dados['porPagina'], dados['total'])


JS_LER_PRODUTOS = """
const h = [...document.querySelectorAll('h6')]
    .find(e => e.textContent.trim() === 'Produtos da categoria');
if (!h) return [];
// A tabela dos produtos já associados é a primeira depois desse título. A outra
// tabela da página fica sob "Seleção Manual de Produtos" (busca no catálogo
// inteiro) e não pode ser confundida com esta.
const tabela = [...document.querySelectorAll('table')].find(
    t => h.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING);
if (!tabela) return [];
return [...tabela.querySelectorAll('tbody tr')].map(tr => {
    const a = tr.querySelector('a[href*="/dashboard/products/"]');
    const tds = tr.querySelectorAll('td');
    if (!a) return null;
    return {
        href: a.getAttribute('href'),
        nome: a.textContent.trim(),
        sku: tds.length > 2 ? tds[2].textContent.trim() : ''
    };
}).filter(Boolean);
"""


def esperar_linhas(driver):
    """Espera a tabela de produtos associados ter linhas e devolve o que leu.

    Espera também o número de linhas parar de crescer: a tabela é preenchida em
    lotes e uma leitura no meio do caminho enxergaria só parte da coleção.
    """
    anterior, estaveis, produtos = -1, 0, []
    fim = time.time() + TEMPO_ESPERA
    while time.time() < fim:
        produtos = driver.execute_script(f'return (function() {{{JS_LER_PRODUTOS}}})();') or []
        if produtos and len(produtos) == anterior:
            estaveis += 1
            if estaveis >= 2:
                return produtos
        else:
            estaveis = 0
        anterior = len(produtos)
        time.sleep(0.4)
    return produtos


def ler_produtos_da_colecao(driver):
    """Devolve {sku_normalizado: {href, nome, sku}} da tabela 'Produtos da categoria'."""
    esperar_linhas(driver)
    ampliar_pagina_de_produtos(driver)
    # Trocar o tamanho da página remonta a tabela: por um instante ela fica sem
    # nenhuma linha, e ler nesse intervalo devolveria uma coleção vazia.
    produtos = esperar_linhas(driver)

    mapa = OrderedDict()
    for p in produtos:
        chave = normalizar_sku(p['sku'])
        if chave and chave not in mapa:
            mapa[chave] = p

    dados = driver.execute_script(f'return (function() {{{JS_SELETOR_POR_PAGINA}}})();') or {}
    total = dados.get('total', 0)
    if total and len(produtos) < total:
        log.warning('A coleção informa %d produtos, mas só %d foram lidos.', total, len(produtos))
    log.info('%d produtos lidos na coleção.', len(mapa))
    return mapa


# ---------------------------------------------------------------- produto

def definir_ordenacao(driver, href, ordem, salvar=True):
    """Abre a página do produto, escreve a ordenação e (se salvar) clica em Atualizar."""
    driver.get(href if href.startswith('http') else BASE + href)
    espera = WebDriverWait(driver, TEMPO_ESPERA)

    campo = espera.until(EC.element_to_be_clickable((By.NAME, 'ordering')))
    driver.execute_script('arguments[0].scrollIntoView({block:"center"});', campo)
    anterior = campo.get_attribute('value')

    campo.click()
    campo.send_keys(Keys.CONTROL, 'a')
    campo.send_keys(Keys.DELETE)
    campo.send_keys(str(ordem))

    atual = campo.get_attribute('value')
    if atual != str(ordem):
        raise RuntimeError(f'campo "Ordenação" ficou "{atual}", esperado "{ordem}"')
    log.info('Ordenação: "%s" -> "%s".', anterior, atual)

    if not salvar:
        # Prova visual do ensaio: o campo preenchido, ainda no centro da tela.
        caminho = os.path.join(PASTA, 'teste_ordenacao.png')
        driver.save_screenshot(caminho)
        log.info('Captura de tela salva em %s', caminho)

    botao = espera.until(EC.presence_of_element_located(
        (By.XPATH, "//button[@type='submit'][normalize-space()='Atualizar']")
    ))
    driver.execute_script('arguments[0].scrollIntoView({block:"center"});', botao)

    if not salvar:
        log.info('MODO TESTE: "Atualizar" localizado, mas NÃO foi clicado.')
        return anterior

    espera.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[@type='submit'][normalize-space()='Atualizar']")
    )).click()
    time.sleep(PAUSA_APOS_SALVAR)
    log.info('Salvo.')
    return anterior


# ---------------------------------------------------------------- execução

def executar(linhas, usuario, senha, salvar=True, manter_aberto=False):
    resultados = []
    driver = abrir_navegador()
    try:
        fazer_login(driver, usuario, senha)

        # Agrupa por coleção para abrir cada carrossel uma vez só.
        por_colecao = OrderedDict()
        for item in linhas:
            por_colecao.setdefault(item['colecao'], []).append(item)

        for colecao, itens in por_colecao.items():
            log.info('=' * 70)
            log.info('COLEÇÃO "%s" — %d produto(s)', colecao, len(itens))
            try:
                abrir_colecao(driver, colecao)
                catalogo = ler_produtos_da_colecao(driver)
            except Exception as erro:
                log.error('Falha ao abrir a coleção "%s": %s', colecao, erro)
                for item in itens:
                    resultados.append({**item, 'status': 'ERRO', 'detalhe': f'coleção: {erro}'})
                continue

            for item in itens:
                produto = catalogo.get(item['sku'])
                if not produto:
                    log.error('Linha %d: SKU %s não está na coleção "%s".',
                              item['linha'], item['sku_planilha'], colecao)
                    resultados.append({**item, 'status': 'NÃO ENCONTRADO', 'detalhe': ''})
                    continue

                log.info('-' * 70)
                log.info('Linha %d | SKU %s (painel: %s) | ordem %s | %s',
                         item['linha'], item['sku_planilha'], produto['sku'],
                         item['ordem'], produto['nome'])
                try:
                    anterior = definir_ordenacao(driver, produto['href'], item['ordem'], salvar)
                    resultados.append({
                        **item,
                        'status': 'OK' if salvar else 'TESTE (não salvo)',
                        'detalhe': f'{produto["nome"]} | ordenação anterior: {anterior}',
                    })
                except Exception as erro:
                    log.error('Linha %d: falha ao ordenar — %s', item['linha'], erro)
                    resultados.append({**item, 'status': 'ERRO', 'detalhe': str(erro)})
    finally:
        if manter_aberto:
            try:
                input('\n>>> Navegador aberto para conferência. Pressione ENTER para fechar... ')
            except EOFError:
                pass    # rodando sem terminal interativo: fecha direto
        driver.quit()
    return resultados


def resumir(resultados):
    log.info('=' * 70)
    for r in resultados:
        log.info('linha %-4s | %-12s | ordem %-4s | %-20s | %s',
                 r['linha'], r['sku_planilha'], r['ordem'], r['colecao'], r['status'])
    ok = sum(1 for r in resultados if r['status'].startswith(('OK', 'TESTE')))
    log.info('%d de %d processados com sucesso.', ok, len(resultados))


def main():
    parser = argparse.ArgumentParser(description='Ordena os produtos dos carrosséis da Sportbay.')
    parser.add_argument('--teste', action='store_true',
                        help='processa só a 1ª linha e NÃO clica em Atualizar')
    parser.add_argument('--linhas', type=int, default=0,
                        help='processa apenas as N primeiras linhas da planilha')
    parser.add_argument('--planilha', default=os.path.join(PASTA, NOME_PLANILHA))
    args = parser.parse_args()

    # .env compartilhado: hoje fica em "Analise de preços site\Analise";
    # a raiz e a pasta deste script continuam valendo como reserva.
    raiz = os.path.dirname(PASTA)
    for pasta_env in (os.path.join(raiz, 'Analise'), raiz, PASTA):
        load_dotenv(os.path.join(pasta_env, '.env'))
    usuario = os.getenv('SPORTBAY_USUARIO')
    senha = os.getenv('SPORTBAY_SENHA')
    if not usuario or not senha:
        raise SystemExit('Defina SPORTBAY_USUARIO e SPORTBAY_SENHA no arquivo .env.')

    linhas = ler_planilha(args.planilha)
    if args.teste:
        linhas = linhas[:1]
    elif args.linhas:
        linhas = linhas[:args.linhas]
    if not linhas:
        raise SystemExit('Nada para processar.')

    log.info('%d linha(s) para processar. Salvar: %s', len(linhas), 'NÃO (teste)' if args.teste else 'sim')
    resultados = executar(linhas, usuario, senha, salvar=not args.teste, manter_aberto=args.teste)
    resumir(resultados)


if __name__ == '__main__':
    main()
