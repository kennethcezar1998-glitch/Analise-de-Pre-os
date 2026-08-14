"""
Bot de automação — Gerenciamento de Produtos em Coleções (Sportbay Marketplace)

Remove todos os produtos atualmente associados a uma coleção e adiciona os
produtos (por shortHash) listados na planilha "ATUALIZAÇÃO DE COLEÇÕES",
excluindo o merchant "Sportbay" da busca de associação manual.

Da planilha o bot usa: Coluna J (nome da coleção alvo, lido da 1ª linha de
dados), Coluna H (shortHash de cada produto) e Coluna A (código usado para
validar o SKU da linha encontrada antes de adicionar).
"""
import os
import re
import time
import logging
import collections
from contextlib import contextmanager

import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
)

# --- CONFIGURAÇÕES DE LOGGING ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('robo_colecoes.log', encoding='utf-8'),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger(__name__)

# --- CONFIGURAÇÕES ---
load_dotenv()

URL_LOGIN = "https://www.marketplace.sportbay.com.br/"
USUARIO = os.getenv("SPORTBAY_USUARIO")
SENHA = os.getenv("SPORTBAY_SENHA")

PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))
NOME_PLANILHA = os.path.join(PASTA_SCRIPT, "ATUALIZAÇÃO DE COLEÇÕES.xlsx")

# Layout da planilha "ATUALIZAÇÃO DE COLEÇÕES" (1 linha de cabeçalho):
#   A Produto Pai | B Código Fabrica.1 | C Código Aux | D Produto | E Categoria
#   F Total | G Link | H Código Site Sportbay | I Menor Preço | J Coleção
COL_CODIGO_IDX = 0      # Coluna A (Produto Pai)
COL_SHORTHASH_IDX = 7   # Coluna H (Código Site Sportbay)
COL_COLECAO_IDX = 9     # Coluna J (Coleção)

TIMEOUT_ELEMENTO = 15
TIMEOUT_LOGIN = 25
MAX_ITERACOES_REMOCAO = 1000

# Tempo que a grade precisa ficar PARADA (mesma contagem/SKU/valor de filtro)
# antes de o bot concluir que o resultado da busca é mesmo um produto
# divergente do código da planilha — ver `_aguardar_resultado_novo`.
ESTABILIDADE_DIVERGENTE = 1.5

# Se True: depois de cada produto, remove o chip de filtro do shortHash antes
# do próximo — dispara uma SEGUNDA consulta ao backend, só pra "resetar" a
# tela, e essa consulta sem filtro renderiza a GRADE SEM FILTRO (todos os
# produtos da loja, primeira linha = um produto qualquer). Era justamente essa
# grade sem filtro que aparecia na tela quando o bot clicava no produto errado.
# Se False (default): pula a limpeza e vai direto pro próximo shortHash — a
# grade sem filtro nunca chega a ser renderizada entre buscas, e o chip do
# shortHash anterior é substituído pelo novo (se o painel acumular chips em vez
# de substituir, `_aguardar_resultado_novo` detecta e remove os antigos). A
# aceitação do resultado em `_aguardar_resultado_novo` (chip do filtro + SKU
# esperado) e a trava atômica de `_clicar_adicionar_validado` garantem que só a
# linha certa seja clicada.
LIMPAR_FILTRO_ENTRE_BUSCAS = False

# Quando a coleção não tem nenhum produto associado, a tabela 'Produtos da
# categoria' não vem vazia: ela renderiza UMA linha de estado vazio — uma
# célula só, com colspan, escrita "Nenhum produto encontrado" — e a paginação
# mostra "0-0 de 0" (HTMLs/NENHUM PRODUTO.txt). Essa linha NÃO é um produto:
# não tem ícone de remover e não deve entrar em contagem nenhuma.
MENSAGEM_TABELA_VAZIA = 'nenhum produto encontrado'

# Seletores da seção 'Produtos da categoria', num lugar só — eram repetidos em
# três funções, e a regra da linha de estado vazio (`not(td[@colspan])`, que
# descarta o "Nenhum produto encontrado") precisa valer igual em todas: se uma
# contasse a linha-mensagem como produto, o total usado para confirmar cada
# adição sairia deslocado em 1.
# Usa //h6[...] (não //*[...]) de propósito: com //*, o texto "Produtos da
# categoria" também "existe" em todo ancestral do heading (div pai, avô, até o
# body, já que `.` pega o texto de toda a subárvore) — isso faz
# `following::table[1]` virar ambíguo (cada ancestral resolve para um "próximo
# table" possivelmente diferente). Travando no <h6> específico, há só um nó de
# contexto e o `following::` sempre aponta pra tabela certa.
_H6_CATEGORIA = "//h6[contains(normalize-space(.),'Produtos da categoria')]"
SELETOR_LINHAS_CATEGORIA = f"{_H6_CATEGORIA}/following::table[1]/tbody/tr[not(td[@colspan])]"
SELETOR_PROXIMA_PAGINA_CATEGORIA = (
    f"{_H6_CATEGORIA}/following::*"
    "[self::a or self::button][contains(@aria-label,'próxima') or contains(@aria-label,'Próxima')"
    " or contains(@class,'next')][not(contains(@class,'disabled'))]"
)

# Mapas para normalizar caixa (inclusive acentuada) via translate() em XPath 1.0,
# que não tem função nativa de comparação case-insensitive.
_MAPA_MINUSCULAS = "abcdefghijklmnopqrstuvwxyzàáâãäåèéêëìíîïòóôõöùúûüçñ"
_MAPA_MAIUSCULAS = "ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÇÑ"

# --- INSTRUMENTAÇÃO DE TEMPOS (diagnóstico de performance) ---
# Liga/desliga a medição por fase sem precisar remover código depois. Com
# True, cada fase relevante (busca, validação de SKU, clique, confirmação de
# contador, limpeza de filtro, remoção) acumula tempo total via
# time.perf_counter(), e um resumo é logado ao final da execução.
MEDIR_TEMPOS = True

_tempos_acumulados = collections.defaultdict(float)
_contagem_fases = collections.defaultdict(int)


@contextmanager
def _medir(fase):
    """Context manager que acumula o tempo gasto no bloco sob `fase`, se
    MEDIR_TEMPOS estiver ligado. Sem overhead perceptível quando desligado
    (só um if antes do yield)."""
    if not MEDIR_TEMPOS:
        yield
        return
    inicio = time.perf_counter()
    try:
        yield
    finally:
        _tempos_acumulados[fase] += time.perf_counter() - inicio
        _contagem_fases[fase] += 1


def _logar_resumo_tempos():
    """Loga o acumulado por fase, ordenado do maior pro menor — chamado uma
    vez ao final de main()."""
    if not MEDIR_TEMPOS or not _tempos_acumulados:
        return
    logger.info("=" * 60)
    logger.info("RESUMO DE TEMPOS POR FASE (MEDIR_TEMPOS=True)")
    total_geral = sum(_tempos_acumulados.values())
    for fase, total in sorted(_tempos_acumulados.items(), key=lambda kv: -kv[1]):
        n = _contagem_fases[fase]
        media = total / n if n else 0
        pct = (total / total_geral * 100) if total_geral else 0
        logger.info(f"  {fase:<24} {total:8.2f}s total | {n:4d}x | média {media:.3f}s | {pct:5.1f}%")
    logger.info(f"  {'TOTAL MEDIDO':<24} {total_geral:8.2f}s")
    logger.info("=" * 60)


# --- HELPERS GENÉRICOS ---
def _tentar_localizar(driver, xpaths, timeout=TIMEOUT_ELEMENTO, clicavel=False):
    """Tenta localizar um elemento usando uma lista de XPaths alternativos,
    em ordem de prioridade. Retorna o primeiro WebElement encontrado ou None."""
    condicao = EC.element_to_be_clickable if clicavel else EC.presence_of_element_located
    for xpath in xpaths:
        try:
            # poll_frequency=0.1 (padrão do Selenium é 0.5s) — quando o
            # elemento aparece logo depois do início da espera, o polling
            # mais fino encurta bastante o atraso até a próxima checagem.
            wait = WebDriverWait(driver, timeout, poll_frequency=0.1)
            elemento = wait.until(condicao((By.XPATH, xpath)))
            return elemento
        except TimeoutException:
            continue
        except Exception as e:
            logger.debug(f"  Falha ao tentar '{xpath}': {e}")
            continue
    return None


def _tentar_clicar(driver, xpaths, timeout=TIMEOUT_ELEMENTO, descricao="elemento"):
    """Localiza (com fallback de seletores) e clica em um elemento. Retorna True/False."""
    elemento = _tentar_localizar(driver, xpaths, timeout=timeout, clicavel=True)
    if elemento is None:
        logger.warning(f"  ⚠ Não encontrou '{descricao}' com nenhum dos seletores tentados")
        return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
        # Sem sleep fixo aqui: `_tentar_localizar` já usou EC.element_to_be_clickable
        # (que exige visível+habilitado) antes de retornar o elemento, então o
        # scrollIntoView só ajusta a posição — não há necessidade de esperar
        # mais nada antes do clique.
        elemento.click()
        return True
    except ElementClickInterceptedException:
        try:
            driver.execute_script("arguments[0].click();", elemento)
            return True
        except Exception as e:
            logger.warning(f"  ⚠ Falha ao clicar em '{descricao}': {e}")
            return False
    except Exception as e:
        logger.warning(f"  ⚠ Falha ao clicar em '{descricao}': {e}")
        return False


def _preencher_campo(driver, elemento, valor):
    """Preenche um input via setter nativo + eventos 'input'/'change' — direto,
    num único execute_script (1 round-trip ao chromedriver), em vez de
    click() + Ctrl+A + Delete + send_keys() caractere-a-caractere +
    get_attribute() (cada um desses é uma chamada HTTP separada ao
    chromedriver). Funciona igual para campos controlados por React/Vue, que
    é justamente por que esse caminho existia como fallback antes — só que
    ele sempre funcionou, então virou o caminho único."""
    driver.execute_script(
        """
        const input = arguments[0];
        const valor = arguments[1];
        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
        setter.call(input, valor);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        elemento,
        valor,
    )


def _debugar_inputs(driver, contexto=""):
    """Loga os atributos de todo <input> visível na página — usado quando um
    seletor esperado não encontra o campo, para ajustar os XPaths sem precisar
    inspecionar o DOM manualmente."""
    try:
        inputs = driver.find_elements(By.TAG_NAME, "input")
        logger.warning(f"  [debug{' - ' + contexto if contexto else ''}] {len(inputs)} <input> encontrado(s) na página:")
        for i, inp in enumerate(inputs):
            try:
                atributos = driver.execute_script(
                    "const el = arguments[0]; const o = {}; "
                    "for (const a of el.attributes) { o[a.name] = a.value; } "
                    "o['_visible'] = !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length); "
                    "return o;",
                    inp,
                )
                logger.warning(f"    input[{i}]: {atributos}")
            except Exception:
                continue
    except Exception as e:
        logger.debug(f"  Falha ao debugar inputs: {e}")


_JS_TRATAR_MODAL_CONFIRMACAO = """
    const seletores = "div[role='dialog'] button, .MuiDialog-root button, [class*='modal'] button";
    for (const botao of document.querySelectorAll(seletores)) {
        const texto = (botao.textContent || '').trim().toLowerCase();
        if (texto !== 'sim' && texto !== 'confirmar' && texto !== 'ok') { continue; }
        if (botao.disabled) { continue; }
        if (!(botao.offsetWidth || botao.offsetHeight || botao.getClientRects().length)) { continue; }
        botao.click();
        return true;
    }
    return false;
"""


def _tratar_modal_confirmacao(driver):
    """Se aparecer um modal de confirmação (ex.: 'Deseja remover?'), clica em confirmar.

    Checagem IMEDIATA e em UMA chamada de JS — o caso comum é NÃO existir modal
    nenhum (a remoção é instantânea ao clicar no ícone), e essa função roda a
    cada produto removido. A versão anterior fazia um `find_elements` com um
    XPath grande e depois `is_displayed()`/`is_enabled()`/`click()` por
    candidato: vários round-trips ao chromedriver por produto só pra concluir
    "não tem modal". Aqui tudo acontece dentro do navegador e volta um booleano.
    """
    try:
        tratou = driver.execute_script(_JS_TRATAR_MODAL_CONFIRMACAO)
    except Exception as e:
        logger.debug(f"  Falha ao checar modal de confirmação: {e}")
        return False
    if tratou:
        logger.debug("  → Modal de confirmação tratado")
        time.sleep(0.5)
    return bool(tratou)


# --- ETAPAS DO FLUXO ---
def login(driver):
    """Realiza login no painel administrativo do Sportbay Marketplace."""
    logger.info("Acessando página de login...")
    driver.get(URL_LOGIN)

    # Atributos reais confirmados no DOM (name="email"/name="password") vêm
    # primeiro — são a busca mais rápida (1 seletor, bate de cara). Os demais
    # ficam como fallback caso o atributo mude, mas cada um só é tentado se
    # o anterior falhar, custando o timeout inteiro — por isso a ordem importa.
    campo_usuario = _tentar_localizar(
        driver,
        [
            "//input[@name='email']",
            "//label[contains(normalize-space(.),'Usuário') or contains(normalize-space(.),'Usuario')]/following::input[1]",
            "//input[@placeholder='Usuário']",
        ],
        timeout=TIMEOUT_ELEMENTO,
    )
    if campo_usuario is None:
        _debugar_inputs(driver, "campo de usuário não encontrado")
        raise RuntimeError("Campo de usuário não encontrado na tela de login")

    campo_senha = _tentar_localizar(
        driver,
        [
            "//input[@name='password']",
            "//input[@type='password']",
            "//input[@placeholder='Senha']",
        ],
        timeout=TIMEOUT_ELEMENTO,
    )
    if campo_senha is None:
        _debugar_inputs(driver, "campo de senha não encontrado")
        raise RuntimeError("Campo de senha não encontrado na tela de login")

    _preencher_campo(driver, campo_usuario, USUARIO)
    _preencher_campo(driver, campo_senha, SENHA)

    clicou = _tentar_clicar(
        driver,
        ["//button[contains(normalize-space(.),'Entrar')]", "//*[normalize-space(text())='Entrar']"],
        descricao="botão Entrar",
    )
    if not clicou:
        raise RuntimeError("Não foi possível clicar no botão 'Entrar'")

    dashboard = _tentar_localizar(
        driver,
        ["//*[contains(normalize-space(.),'E-Commerce')]"],
        timeout=TIMEOUT_LOGIN,
    )
    if dashboard is None:
        raise RuntimeError("Login não concluiu — dashboard 'E-Commerce' não apareceu")

    logger.info("✓ Login realizado com sucesso")


URL_LISTA_CATEGORIAS = "https://www.marketplace.sportbay.com.br/dashboard/categories"


def navegar_para_lista_categorias(driver):
    """Navega até a página 'Lista de Categorias e Coleções' via URL direta
    (mais confiável que clicar no menu lateral, que exige accordion + submenus)."""
    logger.info(f"Navegando direto para: {URL_LISTA_CATEGORIAS}")
    driver.get(URL_LISTA_CATEGORIAS)

    pagina_ok = _tentar_localizar(
        driver,
        ["//*[contains(normalize-space(.),'Coleções')]"],
        timeout=TIMEOUT_ELEMENTO,
    )
    if pagina_ok is None:
        raise RuntimeError("Página 'Lista de Categorias e Coleções' não carregou")

    logger.info("✓ Página de lista de categorias/coleções carregada")


def abrir_colecao(driver, nome_colecao):
    """Localiza (na coluna Coleções) e abre a coleção com o nome exato informado.

    A busca é restrita ao container que fica logo abaixo do <h6>Coleções</h6>
    (irmão direto dele no DOM, confirmado na página real) — não ao restante
    da página. Isso evita bater num item de mesmo nome nas colunas vizinhas
    (Categorias / Categorias dinâmicas), que ficam fora desse container.

    Dentro dele, busca pelo valor de string completo do elemento (`.`, que
    inclui texto de filhos aninhados como <span>/<a>), não apenas `text()`
    (só texto direto), e restringe o resultado ao nó mais específico que
    contém aquele texto — assim funciona em qualquer posição da lista, sem
    depender de rolagem ou de a coleção estar visível na tela no momento da
    busca.
    """
    logger.info(f"Procurando coleção: '{nome_colecao}'")

    coluna_colecoes = "//h6[normalize-space(.)='Coleções']/following-sibling::div[1]"

    cond_exato = f"normalize-space(.)={_xpath_literal(nome_colecao)}"
    cond_ci = (
        f"translate(normalize-space(.), '{_MAPA_MINUSCULAS}', '{_MAPA_MAIUSCULAS}')"
        f"={_xpath_literal(nome_colecao.upper())}"
    )
    seletores_colecao = [
        f"{coluna_colecoes}//*[{cond_exato}][not(.//*[{cond_exato}])]",
        f"{coluna_colecoes}//*[{cond_ci}][not(.//*[{cond_ci}])]",
    ]

    clicou = _tentar_clicar(driver, seletores_colecao, descricao=f"coleção '{nome_colecao}'")
    if not clicou:
        raise RuntimeError(f"Coleção '{nome_colecao}' não encontrada na lista")

    pagina_ok = _tentar_localizar(
        driver,
        ["//*[contains(normalize-space(.),'Editar categoria') or contains(normalize-space(.),'Editar coleção')]"],
        timeout=TIMEOUT_ELEMENTO,
    )
    if pagina_ok is None:
        raise RuntimeError("Página 'Editar categoria ou coleção' não carregou")

    logger.info("✓ Coleção aberta para edição")


def _xpath_literal(texto):
    """Constrói um literal XPath seguro mesmo se o texto contiver aspas simples/duplas."""
    if "'" not in texto:
        return f"'{texto}'"
    if '"' not in texto:
        return f'"{texto}"'
    partes = texto.split("'")
    return "concat('" + "', \"'\", '".join(partes) + "')"


# Classifica o estado atual da tabela 'Produtos da categoria' numa única
# chamada, distinguindo os três casos que importam:
#   'vazia'       → a coleção não tem produto nenhum: a única linha presente é
#                   a de estado vazio ("Nenhum produto encontrado"), que tem
#                   uma célula só. Não há o que remover.
#   'com_produtos'→ há pelo menos uma linha com ícone de remover.
#   'indefinida'  → a tabela ainda não terminou de renderizar (nem mensagem,
#                   nem linha com ícone) — quem chama continua esperando.
# A distinção existe porque "0 linhas" logo depois de abrir a coleção pode ser
# só a tabela ainda carregando; a MENSAGEM é a prova de que o backend
# respondeu e a coleção está mesmo vazia.
_JS_CLASSIFICAR_TABELA_CATEGORIA = """
    const seletorLinhas = arguments[0];
    const seletorCorpo = arguments[1];
    const mensagem = arguments[2];

    const linhas = document.evaluate(
        seletorLinhas, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
    );
    for (let i = 0; i < linhas.snapshotLength; i++) {
        if (linhas.snapshotItem(i).querySelector(
                "[data-testid='RemoveCircleIcon'], [aria-label='Remover produto']")) {
            return 'com_produtos';
        }
    }

    // A linha de estado vazio é descartada por `seletorLinhas`
    // (not(td[@colspan])), então a mensagem é procurada no <tbody> inteiro.
    const corpo = document.evaluate(
        seletorCorpo, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
    ).singleNodeValue;
    if (corpo && (corpo.textContent || '').trim().toLowerCase().includes(mensagem)) {
        return 'vazia';
    }
    return 'indefinida';
"""

_SELETOR_TBODY_CATEGORIA = f"{_H6_CATEGORIA}/following::table[1]/tbody"


def _classificar_tabela_categoria(driver, timeout=TIMEOUT_ELEMENTO):
    """Espera a tabela 'Produtos da categoria' terminar de renderizar e devolve
    'vazia', 'com_produtos' ou 'indefinida' (se estourar o timeout sem os dois
    sinais). Ver `_JS_CLASSIFICAR_TABELA_CATEGORIA`."""
    fim = time.time() + timeout
    while True:
        try:
            situacao = driver.execute_script(
                _JS_CLASSIFICAR_TABELA_CATEGORIA,
                SELETOR_LINHAS_CATEGORIA,
                _SELETOR_TBODY_CATEGORIA,
                MENSAGEM_TABELA_VAZIA,
            )
        except Exception as e:
            logger.debug(f"  Falha ao classificar a tabela 'Produtos da categoria': {e}")
            situacao = 'indefinida'
        if situacao != 'indefinida' or time.time() >= fim:
            return situacao
        time.sleep(0.1)


# Localiza a PRIMEIRA linha da tabela informada, procura o ícone de remover
# DENTRO dela e clica — tudo num único tick de JS. Devolve a própria linha para
# que quem chamou possa esperar a staleness dela (prova de que a remoção
# aconteceu de fato, e não só de que o clique foi disparado).
_JS_REMOVER_PRIMEIRA_LINHA = """
    const seletorLinhas = arguments[0];
    const linhas = document.evaluate(
        seletorLinhas, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
    );
    if (linhas.snapshotLength === 0) { return {clicou: false, linha: null, restantes: 0}; }

    const linha = linhas.snapshotItem(0);
    const alvo = linha.querySelector("[data-testid='RemoveCircleIcon'], [aria-label='Remover produto']");
    if (!alvo) { return {clicou: false, linha: null, restantes: linhas.snapshotLength}; }

    alvo.scrollIntoView({block: 'center'});
    for (const tipo of ['mousedown', 'mouseup', 'click']) {
        alvo.dispatchEvent(new MouseEvent(tipo, {bubbles: true, cancelable: true, view: window}));
    }
    return {clicou: true, linha: linha, restantes: linhas.snapshotLength};
"""


def remover_todos_produtos_atuais(driver):
    """Remove, um a um, todos os produtos já associados à coleção, percorrendo
    todas as páginas da listagem 'Produtos da categoria'.

    Se a coleção já estiver vazia (a tabela mostra 'Nenhum produto encontrado'),
    retorna na hora sem entrar no laço de remoção — o fluxo segue direto para a
    'Seleção Manual de Produtos'. Só esse caso é encurtado: havendo qualquer
    produto na coleção, a remoção acontece normalmente.
    """
    logger.info("Removendo produtos atuais da coleção...")

    secao_produtos = _tentar_localizar(driver, [_H6_CATEGORIA], timeout=TIMEOUT_ELEMENTO)
    if secao_produtos is None:
        raise RuntimeError("Seção 'Produtos da categoria' não encontrada na página de edição")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", secao_produtos)

    # O <h6> aparecer não quer dizer que a tabela já respondeu: logo depois de
    # abrir a coleção ela ainda pode estar sem linha nenhuma por estar
    # carregando. Por isso a decisão espera um sinal conclusivo — a mensagem de
    # vazio OU uma linha com ícone de remover — em vez de olhar só a contagem.
    situacao = _classificar_tabela_categoria(driver)
    if situacao == 'vazia':
        logger.info(
            "✓ Coleção já está sem produtos ('Nenhum produto encontrado') — "
            "nada a remover, indo direto para a Seleção Manual de Produtos"
        )
        return
    if situacao == 'indefinida':
        logger.warning(
            "  ⚠ Não foi possível confirmar se a coleção tem produtos "
            "(nem a mensagem de vazio, nem linha com ícone de remover) — seguindo com a remoção"
        )

    seletor_linhas_categoria = SELETOR_LINHAS_CATEGORIA
    seletor_proxima_pagina = SELETOR_PROXIMA_PAGINA_CATEGORIA

    total_removidos = 0
    for _ in range(MAX_ITERACOES_REMOCAO):
        with _medir('remocao_item'):
            # Localizar + rolar + clicar num único bloco de JS: antes eram 3 a 4
            # round-trips ao chromedriver por produto (WebDriverWait com XPath
            # `following::table[1]`, scrollIntoView, click) mais um
            # `find_elements` com XPath grande só pra checar modal — tudo isso
            # multiplicado por centenas de produtos. Aqui vira 1 chamada, e o
            # ícone é procurado DENTRO da primeira linha (não no escopo da
            # tabela inteira), então nunca há dúvida sobre qual produto sai.
            resposta = driver.execute_script(_JS_REMOVER_PRIMEIRA_LINHA, seletor_linhas_categoria)

            if not resposta.get('clicou'):
                proxima = _tentar_localizar(driver, [seletor_proxima_pagina], timeout=3, clicavel=True)
                if proxima:
                    try:
                        _clicar_proxima_pagina_e_aguardar(driver, proxima, seletor_linhas_categoria, timeout_fallback=1)
                        continue
                    except Exception:
                        break
                break

            # Checagem rápida de modal de confirmação — na maioria dos casos não
            # existe modal (remoção é instantânea ao clicar no ícone).
            _tratar_modal_confirmacao(driver)

            linha_removida = resposta.get('linha')
            try:
                WebDriverWait(driver, 3, poll_frequency=0.05).until(EC.staleness_of(linha_removida))
            except TimeoutException:
                time.sleep(0.3)

            total_removidos += 1
            logger.info(f"  ➖ Produto removido (total até agora: {total_removidos})")

    logger.info(f"✓ Remoção concluída — {total_removidos} produto(s) removido(s)")


def limpar_filtro_merchant(driver):
    """Remove a tag 'Sportbay - 04.200.198/0001-08' pré-selecionada no filtro
    Merchant (MUI Autocomplete) da Seleção Manual de Produtos.

    Trava de segurança: só retorna depois de CONFIRMAR (via nova consulta ao
    DOM) que a tag realmente sumiu da tela. Sem essa confirmação, a lista de
    'Seleção Manual de Produtos' continua filtrada por Sportbay, e o primeiro
    resultado dessa lista (não o produto certo) pode acabar sendo adicionado
    por engano — foi exatamente esse o bug relatado.
    """
    logger.info("Removendo filtro de Merchant 'Sportbay'...")

    seletor_chip_sportbay = "//*[contains(@class,'MuiChip-root')][contains(normalize-space(.),'Sportbay')]"
    seletores_tag_x = [
        # Ícone real observado no DOM: <svg data-testid="CancelIcon"> dentro do chip com texto "Sportbay"
        f"{seletor_chip_sportbay}//*[@data-testid='CancelIcon']",
        "//*[contains(normalize-space(.),'Sportbay')]//*[@data-testid='CancelIcon']",
    ]

    if not _tentar_clicar(driver, seletores_tag_x, descricao="'x' do filtro Merchant Sportbay"):
        raise RuntimeError(
            "Não foi possível clicar no 'x' do filtro Merchant 'Sportbay' — "
            "abortando antes de adicionar produtos para não arriscar adicionar o item errado"
        )

    fim = time.time() + TIMEOUT_ELEMENTO
    while time.time() < fim:
        if not driver.find_elements(By.XPATH, seletor_chip_sportbay):
            logger.info("✓ Filtro de Merchant removido e confirmado (tag não existe mais no DOM)")
            return
        time.sleep(0.3)

    raise RuntimeError(
        "Tag do filtro Merchant 'Sportbay' ainda aparece na tela após tentar remover — "
        "abortando antes de adicionar produtos para não arriscar adicionar o item errado"
    )


def _linha_unica_atual(driver, seletor_linhas):
    """Retorna a única linha atual da tabela de resultados, ou None se a
    tabela tiver 0 ou mais de 1 linha no momento."""
    linhas = driver.find_elements(By.XPATH, seletor_linhas)
    return linhas[0] if len(linhas) == 1 else None


def _esta_stale(referencia):
    """True se `referencia` (WebElement) não existir mais no DOM. `None`
    conta como "já stale" (não havia nada pra ficar velho)."""
    if referencia is None:
        return True
    try:
        referencia.is_enabled()
        return False
    except StaleElementReferenceException:
        return True


def _clicar_proxima_pagina_e_aguardar(driver, proxima, seletor_linhas, timeout_fallback):
    """Clica no botão de próxima página e aguarda a página realmente mudar
    (staleness da primeira linha atual) em vez de um `time.sleep()` fixo.
    Se não houver linha pra rastrear (página vazia) ou a staleness nunca
    ocorrer, cai no fallback de `timeout_fallback` segundos — mesmo teto que
    o sleep fixo que isso substitui, então nunca espera MAIS do que antes,
    só espera MENOS quando a página muda mais rápido que isso."""
    # Nota: não usa `_linha_unica_atual` aqui de propósito — essas tabelas
    # paginadas normalmente têm VÁRIAS linhas por página (não 1), então só
    # precisamos da primeira linha como referência de staleness, não de uma
    # contagem exata.
    linhas = driver.find_elements(By.XPATH, seletor_linhas)
    linha_ref = linhas[0] if linhas else None

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", proxima)
    proxima.click()

    if linha_ref is None:
        time.sleep(timeout_fallback)
        return

    fim = time.time() + timeout_fallback
    while time.time() < fim:
        if _esta_stale(linha_ref):
            return
        time.sleep(0.1)


def _aguardar_resultado_novo(driver, seletor_linhas, referencia_antiga, short_hash, codigo_esperado, timeout=10):
    """Aguarda a tabela de resultados mostrar o resultado DESTA busca, e não o
    da busca ANTERIOR (nem a grade sem filtro).

    O shortHash não aparece em nenhuma coluna visível da tabela
    (Nome/SKU/Situação/Merchant/Criação/Estado), então não dá pra confirmar a
    busca comparando o texto da linha com o hash. As provas usadas são, nesta
    ordem:

      1. O filtro de shortHash APLICADO na tela é exatamente o desta busca —
         hoje isso é o chip 'shortHash: P-XXXX' que o painel cria ao apertar
         Enter (o <input> é esvaziado nesse momento); em telas antigas, o
         próprio valor do campo. Ver `_filtro_confere`. Sem filtro nenhum
         aplicado (grade completa) ou com o filtro de outra busca, nada do que
         a tabela mostra pertence a esta busca.
      2. A tabela tem exatamente 1 linha, já com conteúdo renderizado (o <tr>
         pode entrar no DOM um instante antes das células — aceitar cedo
         demais fazia a validação ler SKU/Merchant vazios e reprovar produto
         certo).
      3. O SKU dessa linha bate com o código esperado da planilha (Coluna A) —
         prova direta de que é o produto certo, independente de qual nó do DOM
         o React reaproveitou. É o caminho normal e o mais rápido.
      4. Se o código da planilha não for utilizável para validar (vazio/nan),
         exige uma prova de que a grade foi RE-RENDERIZADA para esta busca:
         ou o filtro aplicado mudou para o desta busca durante a espera, ou
         `referencia_antiga` (a linha que estava na tela antes) saiu do DOM.

    O critério 3 também é o que permite aceitar rápido sem depender de
    staleness — o React às vezes reaproveita o mesmo <tr>, e nesse caso a
    linha antiga nunca "fica stale" mesmo com a tabela já atualizada.

    Retorna (ok, estado) — `estado` é o `EstadoResultado` da última consulta,
    com linha/sku/merchant/nome já lidos (sem round-trip extra depois).
    """
    fim = time.time() + timeout
    pode_validar_por_sku = bool(_normalizar_codigo(codigo_esperado))

    estado = _consultar_estado_resultado(driver, seletor_linhas, referencia_antiga)
    ref_ja_saiu = estado.ref_stale
    # Se o filtro desta busca AINDA NÃO estava aplicado quando a espera começou
    # e passar a estar no meio dela, isso por si só prova que a grade foi
    # re-renderizada para esta busca — serve como alternativa à staleness da
    # linha antiga (que o React às vezes nunca dispara, por reaproveitar o <tr>).
    filtro_mudou = not _filtro_confere(estado, short_hash)
    proxima_limpeza = 0.0
    assinatura_anterior, estavel_desde = None, time.time()
    while True:
        filtro_ok = _filtro_confere(estado, short_hash)
        # `assentado`: a grade parou numa única linha já renderizada — seja um
        # produto, seja a linha de "Nenhum produto encontrado".
        assentado = filtro_ok and estado.contagem == 1 and estado.tem_conteudo
        pronto = assentado and not estado.vazio

        # Caminho normal e mais rápido: o SKU da linha é o do produto esperado.
        if pronto and pode_validar_por_sku and _codigo_bate_com_sku(codigo_esperado, estado.sku, estado.merchant):
            return True, estado

        # O painel acumulou o chip desta busca com o de uma busca anterior — a
        # grade fica filtrada por dois hashes ao mesmo tempo e não casa com
        # nada. Remove o chip antigo (com trava de tempo pra não ficar clicando
        # a cada iteração do polling) e segue esperando.
        if len(estado.hashes_filtro) > 1 and short_hash in estado.hashes_filtro:
            if time.time() >= proxima_limpeza:
                proxima_limpeza = time.time() + 1.0
                antigos = _limpar_chips_shorthash(driver, manter=short_hash)
                if antigos:
                    logger.debug(f"  Chip(s) de filtro remanescente(s) removido(s): {antigos}")

        # A linha na tela não é a esperada (ou não há como validar por SKU).
        # Só devolve esse estado depois que ele ficar PARADO por
        # ESTABILIDADE_DIVERGENTE — uma linha remanescente da busca anterior é
        # substituída em seguida, então esperar a tela assentar evita reportar
        # "código divergente" (e pular o produto) por causa de um piscar de
        # tela. Um resultado realmente divergente estabiliza na hora e custa
        # só esse tempinho extra.
        # Mesma espera de estabilidade vale para a linha de "Nenhum produto
        # encontrado": ela também pode ser um quadro intermediário enquanto o
        # backend responde. Assentada, é resposta definitiva — devolve
        # `ok=False` (não há produto pra clicar) e quem chama diferencia pelo
        # `estado.vazio`, sem repetir uma busca que já teve resposta.
        assinatura = (estado.contagem, estado.sku, estado.vazio) if assentado else None
        if assinatura != assinatura_anterior:
            assinatura_anterior, estavel_desde = assinatura, time.time()
        if assentado and (ref_ja_saiu or filtro_mudou) and (time.time() - estavel_desde) >= ESTABILIDADE_DIVERGENTE:
            return (not estado.vazio), estado

        if time.time() >= fim:
            return False, estado
        time.sleep(0.05)
        # Depois que a referência antiga sai do DOM, para de mandá-la ao script
        # (uma referência já stale só geraria exceção a cada iteração).
        estado = _consultar_estado_resultado(
            driver, seletor_linhas, None if ref_ja_saiu else referencia_antiga
        )
        ref_ja_saiu = ref_ja_saiu or estado.ref_stale
        filtro_mudou = filtro_mudou and _filtro_confere(estado, short_hash)


EstadoResultado = collections.namedtuple(
    'EstadoResultado',
    'ref_stale contagem tem_conteudo linha sku merchant nome campo_valor hashes_filtro vazio',
)

# O painel NÃO deixa mais o shortHash pesquisado dentro do <input>: ao apertar
# Enter ele transforma o filtro num chip ("shortHash: P-XXXXXXXXXX") logo acima
# da grade e ESVAZIA o campo de texto — confirmado no DOM real
# (HTMLs/PRODUTO PESQUISADO PELO SHOTHASH.txt). Era isso que quebrava o bot:
# a prova nº 1 de `_aguardar_resultado_novo` comparava `input.value` com o hash
# da busca, que passou a ser sempre '', então nenhum resultado era aceito (a
# linha certa aparecia na tela, o bot estourava os 10s e pulava o produto).
#
# Este helper lê os chips de filtro aplicados. Continua devolvendo também o
# valor do <input> porque as duas formas convivem: se o painel voltar a manter
# o texto no campo (ou numa tela em que o chip ainda não exista), a validação
# cai de volta no valor do campo — ver `_filtro_confere`.
_JS_HELPERS_FILTRO = """
    function hashesFiltroAplicados() {
        const encontrados = [];
        for (const rotulo of document.querySelectorAll('.MuiChip-label')) {
            const m = /^shortHash\\s*:\\s*(.+)$/i.exec((rotulo.textContent || '').trim());
            if (m) { encontrados.push(m[1].trim()); }
        }
        return encontrados;
    }
"""

# Colunas da tabela, em ordem: [checkbox, Nome, SKU, Situação, Merchant,
# Criação, Estado, ação] — confirmado no DOM real da 'Seleção Manual de
# Produtos' (HTMLs/Correção do bug.txt).
_JS_CONSULTAR_ESTADO_RESULTADO = _JS_HELPERS_FILTRO + """
    const seletor = arguments[0];
    const refAntiga = arguments[1];
    const mensagemVazia = arguments[2];
    const resultado = document.evaluate(
        seletor, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
    );
    const contagem = resultado.snapshotLength;
    let refStale = true;
    if (refAntiga) {
        try { refStale = !refAntiga.isConnected; } catch (e) { refStale = true; }
    }
    const campo = document.querySelector("input[placeholder*='shortHash']");
    let temConteudo = false, linha = null, sku = '', merchant = '', nome = '', vazio = false;
    if (contagem === 1) {
        const tr = resultado.snapshotItem(0);
        const texto = (tr.textContent || '').trim();
        temConteudo = texto !== '';
        const tds = tr.getElementsByTagName('td');
        // Mesma linha de estado vazio da tabela de categoria: célula única com
        // colspan e a mensagem do painel. É uma linha "cheia" (tem texto), mas
        // não é produto nenhum — sem isso, o SKU sairia '' e um shortHash
        // inexistente no site era reportado como "código divergente".
        vazio = tds.length <= 1 && texto.toLowerCase().includes(mensagemVazia);
        if (temConteudo && !vazio) {
            nome = tds[1] ? tds[1].textContent.trim() : '';
            sku = tds[2] ? tds[2].textContent.trim() : '';
            merchant = tds[4] ? tds[4].textContent.trim() : '';
            linha = tr;
        }
    }
    return {
        refStale: refStale, contagem: contagem, temConteudo: temConteudo, linha: linha,
        sku: sku, merchant: merchant, nome: nome, vazio: vazio,
        campoValor: campo ? campo.value.trim() : null,
        hashesFiltro: hashesFiltroAplicados(),
    };
"""

# Remove os chips de filtro por shortHash, exceto o de `manter` (quando
# informado). Só é usado se o painel ACUMULAR chips a cada Enter em vez de
# substituir o anterior — nesse caso duas buscas seguidas viram um filtro
# "shortHash A E shortHash B", que não casa com produto nenhum. Remove UM chip
# por chamada de propósito: cada remoção re-renderiza a lista inteira de chips,
# então os nós coletados antes ficariam obsoletos no meio do laço.
_JS_REMOVER_CHIP_SHORTHASH = _JS_HELPERS_FILTRO + """
    const manter = arguments[0];
    for (const chip of document.querySelectorAll('.MuiChip-root')) {
        const rotulo = chip.querySelector('.MuiChip-label');
        if (!rotulo) { continue; }
        const m = /^shortHash\\s*:\\s*(.+)$/i.exec((rotulo.textContent || '').trim());
        if (!m) { continue; }
        if (manter !== null && m[1].trim() === manter) { continue; }
        const x = chip.querySelector("[data-testid='CancelIcon']");
        if (!x) { continue; }
        for (const tipo of ['mousedown', 'mouseup', 'click']) {
            x.dispatchEvent(new MouseEvent(tipo, {bubbles: true, cancelable: true, view: window}));
        }
        return m[1].trim();
    }
    return null;
"""


def _remover_chip_shorthash(driver, manter=None):
    """Remove um chip de filtro de shortHash (o primeiro que não seja `manter`).
    Retorna o hash removido, ou None se não havia nenhum a remover."""
    try:
        return driver.execute_script(_JS_REMOVER_CHIP_SHORTHASH, manter)
    except Exception as e:
        logger.debug(f"  Falha ao remover chip de shortHash: {e}")
        return None


def _limpar_chips_shorthash(driver, manter=None, max_remocoes=5):
    """Remove todos os chips de filtro de shortHash (menos `manter`), um a um.

    Para assim que uma remoção não faz efeito: se o mesmo chip reaparece na
    chamada seguinte, o clique no 'x' não removeu nada e insistir só gastaria
    round-trips (e faria a função reportar como removido algo que continua na
    tela). Retorna a lista dos hashes efetivamente removidos."""
    removidos = []
    ultimo = None
    for _ in range(max_remocoes):
        hash_removido = _remover_chip_shorthash(driver, manter)
        if hash_removido is None:
            break
        if hash_removido == ultimo:
            logger.warning(f"  ⚠ Chip de filtro '{hash_removido}' não saiu da tela após clicar no 'x'")
            removidos.pop()
            break
        removidos.append(hash_removido)
        ultimo = hash_removido
    return removidos


def _filtro_confere(estado, short_hash):
    """True se o filtro de shortHash aplicado na tela for EXATAMENTE o desta
    busca. Prioriza o chip (comportamento atual do painel); se não houver chip
    nenhum, cai no valor do <input> (comportamento antigo). Um chip a mais
    (filtro de outra busca ainda ativo) reprova — a grade estaria filtrada por
    dois hashes ao mesmo tempo."""
    if estado.hashes_filtro:
        return estado.hashes_filtro == [short_hash]
    return estado.campo_valor == short_hash


def _consultar_estado_resultado(driver, seletor_linhas, referencia_antiga=None):
    """Uma única chamada ao navegador (via `document.evaluate`) que retorna,
    de uma vez: se `referencia_antiga` já saiu do DOM (`.isConnected`), a
    contagem atual de linhas, se a linha (quando há exatamente 1) já tem
    conteúdo renderizado, a própria linha, o SKU/Merchant/Nome dela e o valor
    atual do campo de shortHash. Usado no laço de polling de
    `_aguardar_resultado_novo`: colapsa em 1 round-trip o que seriam várias
    chamadas separadas por iteração, e já entrega o SKU/Merchant que antes
    exigiam mais um round-trip (`_extrair_sku_e_merchant`) depois da espera.
    """
    try:
        dados = driver.execute_script(
            _JS_CONSULTAR_ESTADO_RESULTADO, seletor_linhas, referencia_antiga, MENSAGEM_TABELA_VAZIA
        )
    except StaleElementReferenceException:
        # referencia_antiga não pôde nem ser resolvida como argumento do
        # script — equivale a já estar stale; refaz a consulta sem ela.
        dados = driver.execute_script(
            _JS_CONSULTAR_ESTADO_RESULTADO, seletor_linhas, None, MENSAGEM_TABELA_VAZIA
        )
        dados['refStale'] = True
    return EstadoResultado(
        dados['refStale'], dados['contagem'], dados['temConteudo'], dados['linha'],
        dados['sku'], dados['merchant'], dados['nome'], dados['campoValor'],
        dados['hashesFiltro'], dados['vazio'],
    )


# Revalida e clica NUM ÚNICO bloco de JS, sem devolver o controle ao Python
# entre a checagem e o clique. Esse é o ponto exato do bug: antes, o bot
# validava o SKU da linha encontrada e só DEPOIS procurava o ícone de
# adicionar com um XPath de escopo de TABELA INTEIRA
# (`...//*[@data-testid='AddBoxIcon']`, que devolve o PRIMEIRO ícone da
# tabela). Se a grade re-renderizasse nesse intervalo — resposta atrasada da
# busca de limpeza, que traz a grade SEM FILTRO — o primeiro ícone passava a
# ser o da primeira linha da grade completa, e era esse produto (ex.: SKU
# 127101, "Kit Plastico ... Aro 16") que entrava na coleção. O contador ainda
# subia 1, então o clique errado era registrado como sucesso.
#
# Aqui: contagem, SKU, Merchant e valor do campo de filtro são reconferidos e
# o clique acontece no mesmo tick de JS (single-thread, o React não consegue
# re-renderizar no meio), e o ícone é buscado DENTRO da linha validada
# (`linha.querySelector`), nunca no escopo da tabela.
_JS_ADICIONAR_LINHA_VALIDADA = _JS_HELPERS_FILTRO + """
    const seletor = arguments[0];
    const skuValidado = arguments[1];
    const merchantValidado = arguments[2];
    const hashEsperado = arguments[3];

    const resultado = document.evaluate(
        seletor, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
    );
    if (resultado.snapshotLength !== 1) {
        return {ok: false, motivo: 'a grade deixou de ter 1 resultado (' + resultado.snapshotLength + ' linha(s))'};
    }

    // Mesma prova de filtro de `_filtro_confere`: o chip 'shortHash: ...' é a
    // fonte atual da verdade (o <input> é esvaziado assim que a busca dispara);
    // sem chip nenhum, cai no valor do campo.
    const hashes = hashesFiltroAplicados();
    const campo = document.querySelector("input[placeholder*='shortHash']");
    const valorCampo = campo ? campo.value.trim() : null;
    const filtroOk = hashes.length
        ? (hashes.length === 1 && hashes[0] === hashEsperado)
        : (valorCampo === hashEsperado);
    if (!filtroOk) {
        return {ok: false, motivo: "o filtro de shortHash mudou (chips: [" + hashes.join(', ')
            + "], campo: '" + valorCampo + "')"};
    }

    const linha = resultado.snapshotItem(0);
    const tds = linha.getElementsByTagName('td');
    const sku = tds[2] ? tds[2].textContent.trim() : '';
    const merchant = tds[4] ? tds[4].textContent.trim() : '';
    if (sku !== skuValidado || merchant !== merchantValidado) {
        return {ok: false, motivo: "a linha mudou de conteúdo (SKU '" + sku + "', Merchant '" + merchant + "')"};
    }

    const alvo = linha.querySelector("[data-testid='AddBoxIcon'], [aria-label='Adicionar produto']");
    if (!alvo) {
        return {ok: false, motivo: 'ícone de adicionar não encontrado dentro da linha validada'};
    }

    alvo.scrollIntoView({block: 'center'});
    for (const tipo of ['mousedown', 'mouseup', 'click']) {
        alvo.dispatchEvent(new MouseEvent(tipo, {bubbles: true, cancelable: true, view: window}));
    }
    return {ok: true, motivo: '', nome: tds[1] ? tds[1].textContent.trim() : '', sku: sku};
"""


def _clicar_adicionar_validado(driver, seletor_linhas, short_hash, estado):
    """Reconfere o estado da grade e clica no ícone de adicionar DA LINHA
    VALIDADA, tudo em um único bloco de JS (ver comentário acima). Retorna
    (ok, motivo) — em caso de recusa, `motivo` explica o que mudou na tela e
    quem chama refaz a busca em vez de clicar às cegas."""
    try:
        resposta = driver.execute_script(
            _JS_ADICIONAR_LINHA_VALIDADA, seletor_linhas, estado.sku, estado.merchant, short_hash
        )
    except Exception as e:
        # Qualquer falha aqui é tratada como "não clicou" — quem chama refaz a
        # busca. Nunca cair num caminho de clique alternativo sem reconferência.
        return False, f"erro ao executar a checagem de clique: {e}"
    return bool(resposta.get('ok')), resposta.get('motivo') or ''


_cache_campo_filtro = None


def _preencher_e_buscar(driver, seletores_campo_filtro, short_hash):
    """Preenche o campo 'Filtrar por shortHash' e dispara a busca (Enter).

    Reusa o WebElement do campo entre produtos: ele não é re-renderizado a cada
    busca, então relocalizá-lo por XPath a cada shortHash é round-trip puro
    (e o laço roda centenas de vezes). Se o React tiver trocado o nó, o uso
    levanta StaleElementReferenceException e o campo é relocalizado uma vez.

    Retorna False se o campo não existir mais na página (erro grave)."""
    global _cache_campo_filtro
    for _ in (1, 2):
        if _cache_campo_filtro is None:
            _cache_campo_filtro = _tentar_localizar(driver, seletores_campo_filtro, timeout=TIMEOUT_ELEMENTO)
            if _cache_campo_filtro is None:
                return False
        try:
            _preencher_campo(driver, _cache_campo_filtro, short_hash)
            # A busca é disparada ao apertar Enter, não ao digitar. Enter real
            # (send_keys) de propósito: o campo fica dentro de um <form>, e o
            # painel depende do submit implícito — um KeyboardEvent sintético
            # não o dispara.
            _cache_campo_filtro.send_keys(Keys.ENTER)
            return True
        except StaleElementReferenceException:
            _cache_campo_filtro = None
    return False


def _aguardar_linha_sumir(referencia, timeout=10):
    """Aguarda a linha `referencia` (o resultado do shortHash recém
    processado) sair do DOM antes de liberar a próxima busca — via staleness,
    pelo mesmo motivo de `_aguardar_resultado_novo` (shortHash não aparece no
    texto da linha, então não dá pra confirmar por conteúdo). Sem essa
    espera, limpar o campo e já disparar o Enter seguinte pode fazer o SPA
    cancelar/perder a requisição de busca anterior no meio do caminho."""
    if referencia is None:
        return True
    fim = time.time() + timeout
    while time.time() < fim:
        if _esta_stale(referencia):
            return True
        time.sleep(0.1)
    return False


def _normalizar_codigo(codigo_planilha):
    """Normaliza o código da Coluna A da planilha para comparação com o SKU:
    descarta o artefato '.0' (leitura como float pelo pandas) e o prefixo
    'PAI-', devolvendo em caixa alta. Retorna '' quando não há código
    utilizável (vazio/nan) — usado também por `_aguardar_resultado_novo` para
    saber se dá pra confirmar o resultado da busca pelo SKU."""
    codigo = str(codigo_planilha).strip()
    if codigo.endswith('.0'):  # artefato de leitura como float (ex.: pandas)
        codigo = codigo[:-2]
    if not codigo or codigo.lower() == 'nan':
        return ''
    codigo = codigo.upper()
    if codigo.startswith('PAI-'):
        codigo = codigo[len('PAI-'):]
    return codigo


def _codigo_bate_com_sku(codigo_planilha, sku_site, merchant_site):
    """Compara o código da Coluna A da planilha com o SKU mostrado na linha.

    Normalizações:
      - O prefixo 'PAI-' é descartado dos DOIS lados antes de comparar. Na
        planilha "ATUALIZAÇÃO DE COLEÇÕES" a Coluna A é 'Produto Pai', que
        pode vir com o prefixo (ex.: 'PAI-122361') ou só com o código cru
        (ex.: '122361') — os dois formatos batem com o SKU do site.
      - Se o código da planilha terminar com 'A' e o SKU do site (já sem o
        prefixo) bater com o código SEM o 'A', só é considerado válido se o
        Merchant da linha for 'Sportbay' — fora esse caso o sufixo importa.
    """
    codigo_up = _normalizar_codigo(codigo_planilha)
    if not codigo_up:
        return True  # sem código pra validar, não bloqueia

    sku = sku_site.strip().upper()
    if sku.startswith('PAI-'):
        sku = sku[len('PAI-'):]

    if sku == codigo_up:
        return True

    if codigo_up.endswith('A') and sku == codigo_up[:-1]:
        return merchant_site.strip().upper() == 'SPORTBAY'

    return False


def _ler_total_produtos_categoria(driver, elemento_cache=None):
    """Lê o total real de produtos já associados à coleção a partir do texto
    de paginação da seção 'Produtos da categoria' (ex.: '1-10 de 53' → 53).

    `elemento_cache`, se informado (WebElement já resolvido numa chamada
    anterior), é tentado PRIMEIRO — evita reavaliar o XPath `following::`
    (caro: percorre boa parte do documento) a cada chamada. Isso importa
    porque essa função roda dentro do laço de confirmação de contagem em
    `_tentar_adicionar_um`, que pode chamá-la dezenas de vezes seguidas
    esperando o total incrementar.

    Retorna (total, elemento) — `elemento` é o cache pra reusar na próxima
    chamada (relocalizado automaticamente se o anterior tiver ficado stale).
    `total` é None se não conseguir localizar/interpretar (o chamador tem
    fallback via `_contar_paginando`)."""
    elemento = elemento_cache
    if elemento is not None:
        try:
            texto = driver.execute_script("return arguments[0].textContent;", elemento) or ""
            match = re.search(r'de\s+(\d+)', texto)
            return (int(match.group(1)) if match else None), elemento
        except StaleElementReferenceException:
            elemento = None  # cache perdeu validade (React re-renderizou o nó) — relocaliza abaixo

    seletor_paginacao = f"{_H6_CATEGORIA}/following::div[contains(@class,'MuiTablePagination-root')][1]"
    elemento = _tentar_localizar(driver, [seletor_paginacao], timeout=TIMEOUT_ELEMENTO)
    if elemento is None:
        return None, None
    texto = driver.execute_script("return arguments[0].textContent;", elemento) or ""
    match = re.search(r'de\s+(\d+)', texto)
    return (int(match.group(1)) if match else None), elemento


def _contar_paginando(driver, seletor_linhas, seletor_proxima_pagina, max_paginas=200):
    """Fallback para quando o texto de paginação não pode ser lido: percorre
    todas as páginas da tabela somando a quantidade de linhas. Mais lento que
    ler o contador, por isso só é usado quando este falha."""
    total = 0
    for _ in range(max_paginas):
        total += len(driver.find_elements(By.XPATH, seletor_linhas))
        proxima = _tentar_localizar(driver, [seletor_proxima_pagina], timeout=2, clicavel=True)
        if not proxima:
            break
        try:
            _clicar_proxima_pagina_e_aguardar(driver, proxima, seletor_linhas, timeout_fallback=0.8)
        except Exception:
            break
    return total


def _listar_shorthashes_faltando(driver, lista_shorthashes, seletor_linhas, seletor_proxima_pagina, max_paginas=200):
    """Pagina a tabela 'Produtos da categoria' coletando o texto de todas as
    linhas e verifica quais shortHashes da planilha não aparecem em nenhuma
    delas. Best-effort: o shortHash normalmente NÃO é uma coluna visível na
    tabela, então se NENHUM shortHash bater em NENHUMA linha (sinal de que o
    hash não aparece no texto renderizado), a comparação é considerada
    inconclusiva e retorna lista vazia em vez de reportar todos como
    faltando (o que seria enganoso)."""
    textos_linhas = []
    for _ in range(max_paginas):
        linhas = driver.find_elements(By.XPATH, seletor_linhas)
        for linha in linhas:
            try:
                texto = driver.execute_script("return arguments[0].textContent;", linha) or ""
            except StaleElementReferenceException:
                continue
            textos_linhas.append(texto.lower())
        proxima = _tentar_localizar(driver, [seletor_proxima_pagina], timeout=2, clicavel=True)
        if not proxima:
            break
        try:
            _clicar_proxima_pagina_e_aguardar(driver, proxima, seletor_linhas, timeout_fallback=0.8)
        except Exception:
            break

    if not textos_linhas:
        return []

    texto_completo = " ".join(textos_linhas)
    faltando = [h for h in lista_shorthashes if h.lower() not in texto_completo]

    if len(faltando) == len(lista_shorthashes):
        logger.debug(
            "  Nenhum shortHash apareceu no texto das linhas de 'Produtos da categoria' — "
            "comparação por hash não é confiável aqui, ignorando"
        )
        return []

    return faltando


def _tentar_adicionar_um(driver, short_hash, codigo_esperado, seletores, total_atual, referencia_anterior=None):
    """Tenta adicionar um único shortHash à coleção. Retorna (status, novo_total_atual, linha_nova).

    status é um destes:
      - 'adicionado'      → clique confirmado (total real incrementou em 1)
      - 'sem_efeito'      → clique aconteceu mas o total real não mudou
      - 'codigo_nao_bate' → o SKU da linha encontrada não corresponde ao
                             código da Coluna A da planilha (2ª trava, feita
                             ANTES de clicar — nunca chega a clicar)
      - 'nao_encontrado'  → busca não convergiu para a linha certa, ou botão
                             de adicionar não apareceu
      - 'campo_ausente'   → campo de busca sumiu da página (erro grave, quem
                             chama deve abortar o laço inteiro)

    `total_atual` é o total conhecido de produtos em 'Produtos da categoria'
    antes desta tentativa. Se for None (não foi possível ler o contador nem
    pelo fallback paginado), a verificação de incremento é pulada — o clique
    é aceito só por não ter lançado exceção, com aviso no log.

    `referencia_anterior` é o WebElement da linha do PRODUTO ANTERIOR (o
    `linha_nova` retornado pela chamada anterior), usado só quando
    LIMPAR_FILTRO_ENTRE_BUSCAS=False: como a limpeza de filtro é pulada, a
    tela ainda mostra o resultado anterior no momento em que esta busca
    começa, e é essa referência (não uma reconsulta ao DOM) que
    `_aguardar_resultado_novo` usa como ponto de comparação de staleness na
    PRIMEIRA tentativa. Sem isso, a trava perde a referência certa e volta o
    risco de aceitar o resultado antigo como se fosse novo.

    `linha_nova` (3º item do retorno) é a linha aceita nesta chamada (ou
    None) — quem chama deve passá-la como `referencia_anterior` na PRÓXIMA
    chamada quando LIMPAR_FILTRO_ENTRE_BUSCAS=False.
    """
    seletor_linhas_resultado = seletores["linhas_resultado"]
    seletores_campo_filtro = seletores["campo_filtro"]

    status = None
    novo_total = total_atual
    linha_nova = None

    if not _normalizar_codigo(codigo_esperado):
        logger.warning(
            f"  ⚠ shortHash '{short_hash}': linha sem código na Coluna A da planilha — "
            f"não dá pra validar o SKU do resultado, o produto será aceito só pela busca"
        )

    # Busca + validação + clique, com 1 retentativa: a mesma retentativa cobre
    # tanto "a tabela não convergiu para o resultado desta busca" (ex.: Enter
    # não disparou a busca a tempo) quanto "a grade mudou entre validar e
    # clicar" — nos dois casos refazer a busca do zero é o certo, e nunca se
    # clica num estado de tela que não foi conferido.
    for tentativa_busca in (1, 2):
        # Referência do que está na tela ANTES de disparar esta busca — usada
        # por `_aguardar_resultado_novo` como critério de fallback (staleness)
        # quando não dá pra confirmar o resultado pelo SKU. Na 1ª tentativa,
        # se a limpeza entre buscas está desligada, usa a referência já
        # conhecida (evita uma consulta ao DOM e é o ponto correto de
        # comparação, já que a tela não foi limpa desde a busca anterior).
        # Na 2ª tentativa (retry) sempre reconsulta — o estado pode ter
        # mudado depois da 1ª tentativa falhar.
        if tentativa_busca == 1 and not LIMPAR_FILTRO_ENTRE_BUSCAS and referencia_anterior is not None:
            referencia_antiga = referencia_anterior
        else:
            # Na retentativa, descarta o filtro aplicado antes de repetir: se o
            # chip desta busca já estiver na tela, digitar o mesmo hash de novo
            # pode não mudar nada e a 2ª tentativa viraria cópia da 1ª.
            _limpar_chips_shorthash(driver)
            referencia_antiga = _linha_unica_atual(driver, seletor_linhas_resultado)

        with _medir('busca'):
            if not _preencher_e_buscar(driver, seletores_campo_filtro, short_hash):
                logger.error("  ✗ Campo 'Filtrar por shortHash' não encontrado — abortando adição")
                return 'campo_ausente', total_atual, None

            resultado_ok, estado = _aguardar_resultado_novo(
                driver, seletor_linhas_resultado, referencia_antiga, short_hash, codigo_esperado, timeout=10
            )

        if not resultado_ok:
            # Resposta definitiva do site: a busca rodou (filtro aplicado é o
            # desta linha) e não existe produto com esse shortHash. Repetir a
            # mesma busca daria a mesma coisa, então não gasta a retentativa.
            if estado.vazio:
                logger.warning(
                    f"  ⚠ shortHash '{short_hash}': o site não retornou nenhum produto para esse hash "
                    f"('Nenhum produto encontrado') — pulando"
                )
                status = 'nao_encontrado'
                break

            diagnostico = (
                f"contagem no timeout: {estado.contagem}, filtro aplicado: {estado.hashes_filtro or 'nenhum'}, "
                f"campo: '{estado.campo_valor}'"
            )
            if tentativa_busca == 1:
                logger.warning(
                    f"  ⚠ shortHash '{short_hash}': tabela não convergiu para o resultado desta busca "
                    f"({diagnostico}) — pesquisando novamente..."
                )
                continue
            logger.warning(
                f"  ⚠ shortHash '{short_hash}': tabela não convergiu para o resultado desta busca mesmo após "
                f"repetir ({diagnostico}) — pulando"
            )
            status = 'nao_encontrado'
            break

        linha_nova = estado.linha

        # Trava ANTES de clicar: confere se o SKU da linha realmente
        # corresponde ao código da planilha (Coluna A), tolerando o prefixo
        # "PAI-" e, só para Merchant "Sportbay", a ausência do sufixo "A".
        # SKU/Merchant já vieram na própria consulta de estado da espera — não
        # custa round-trip nenhum aqui.
        with _medir('validacao_sku'):
            codigo_bate = _codigo_bate_com_sku(codigo_esperado, estado.sku, estado.merchant)

        if not codigo_bate:
            logger.warning(
                f"  ⚠ shortHash '{short_hash}': código '{codigo_esperado}' não bate com o SKU da linha "
                f"('{estado.sku}', Merchant: '{estado.merchant}') — pulando para não arriscar adicionar o produto errado"
            )
            status = 'codigo_nao_bate'
            break

        # Reconferência + clique atômicos, dentro da linha validada (ver
        # comentário de `_JS_ADICIONAR_LINHA_VALIDADA`).
        with _medir('clique'):
            clicou, motivo = _clicar_adicionar_validado(driver, seletor_linhas_resultado, short_hash, estado)

        if not clicou:
            if tentativa_busca == 1:
                logger.warning(
                    f"  ⚠ shortHash '{short_hash}': a grade mudou entre validar e clicar ({motivo}) — "
                    f"refazendo a busca em vez de clicar (evita adicionar o produto errado)"
                )
                continue
            logger.warning(
                f"  ⚠ shortHash '{short_hash}': não foi possível clicar com segurança ({motivo}) — pulando"
            )
            status = 'nao_encontrado'
            break

        # Clique disparado na linha certa — verificar efeito real no contador.
        with _medir('confirmacao_contador'):
            if total_atual is None:
                logger.info(
                    f"  ➕ Produto adicionado: {short_hash} — {estado.nome} "
                    f"(SKU {estado.sku}, sem verificação de contagem)"
                )
                status = 'adicionado'
            else:
                fim = time.time() + 8
                confirmado = False
                cache_paginacao = None
                while True:
                    lido, cache_paginacao = _ler_total_produtos_categoria(driver, cache_paginacao)
                    if lido is not None and lido == total_atual + 1:
                        novo_total = lido
                        confirmado = True
                        break
                    if time.time() >= fim:
                        break
                    # 0.05s (em vez de 0.15s): a leitura do contador é uma
                    # única chamada barata sobre um nó já cacheado, então
                    # granularidade mais fina custa quase nada e devolve o
                    # controle assim que o total sobe — isso roda uma vez por
                    # produto e era a 3ª fase mais cara da execução.
                    time.sleep(0.05)
                if confirmado:
                    logger.info(
                        f"  ➕ Produto adicionado: {short_hash} — {estado.nome} (SKU {estado.sku}) "
                        f"(total {total_atual} → {novo_total})"
                    )
                    status = 'adicionado'
                else:
                    logger.warning(
                        f"  ⚠ shortHash '{short_hash}': clique sem efeito "
                        f"(total de produtos continuou em {total_atual})"
                    )
                    status = 'sem_efeito'
        break

    # Descarta o filtro aplicado antes da próxima busca — só depois de tudo
    # acima (confirmação ou desistência), nunca em paralelo com uma requisição
    # de adição ainda em andamento no backend. Remover o chip (e não digitar
    # vazio + Enter) é o que "limpa" a busca hoje: o painel aplica o filtro como
    # chip e já deixa o <input> vazio.
    # Com LIMPAR_FILTRO_ENTRE_BUSCAS=False, pula essa limpeza inteira: o chip do
    # próximo shortHash substitui o atual, e a aceitação de
    # `_aguardar_resultado_novo` (chip do filtro + SKU esperado, ou staleness de
    # `linha_nova` passada como `referencia_anterior` na próxima chamada)
    # continua impedindo que o resultado antigo seja aceito como novo.
    if LIMPAR_FILTRO_ENTRE_BUSCAS:
        with _medir('limpeza_filtro'):
            try:
                if _limpar_chips_shorthash(driver) and not _aguardar_linha_sumir(linha_nova, timeout=10):
                    logger.warning(
                        f"  ⚠ Filtro pode não ter limpado completamente após '{short_hash}' — seguindo mesmo assim"
                    )
            except Exception:
                pass
            # Sem sleep fixo aqui: `_aguardar_linha_sumir` já é espera condicional
            # (staleness) — um sleep(0.5) fixo depois dela era redundante.

    return status, novo_total, linha_nova


def adicionar_produtos(driver, lista_produtos):
    """Adiciona cada produto à coleção via campo 'Filtrar por shortHash'.

    `lista_produtos` é uma lista de tuplas (short_hash, codigo), onde codigo
    vem da Coluna A ('Produto Pai') da planilha e é usado para validar o SKU
    da linha encontrada antes de clicar em adicionar (ver `_codigo_bate_com_sku`).

    Retorna um dict {"adicionados": [...], "nao_encontrados": [...],
    "sem_efeito": [...], "codigo_nao_bate": [...]} com os shortHashes em cada
    categoria (não apenas contagens), para permitir reconciliação/retry por
    quem chamar.
    """
    logger.info(f"Adicionando {len(lista_produtos)} produto(s) por shortHash...")
    codigo_por_hash = {str(h).strip(): c for h, c in lista_produtos}

    # Escopo restrito à tabela de resultados da "Seleção Manual de Produtos"
    # (não a tabela de "Produtos da categoria", que fica acima na mesma página).
    # Usa //h6[...] (não //*[...]) para evitar ambiguidade no following::table[1]
    # — ver comentário equivalente em remover_todos_produtos_atuais.
    seletor_tabela = "//h6[contains(normalize-space(.),'Seleção Manual de Produtos')]/following::table[1]"
    # Não há mais seletor de "botão adicionar" com escopo de tabela: o ícone
    # (<svg data-testid="AddBoxIcon" aria-label="Adicionar produto">) é
    # procurado DENTRO da linha já validada, dentro do mesmo bloco de JS que
    # reconfere o estado da grade — ver `_JS_ADICIONAR_LINHA_VALIDADA`. Um
    # seletor de tabela devolvia o primeiro ícone da grade, que passava a ser
    # de outro produto se a grade re-renderizasse antes do clique.
    seletores = {
        "linhas_resultado": f"{seletor_tabela}/tbody/tr",
        "campo_filtro": ["//input[contains(@placeholder,'shortHash')]"],
    }

    # Sobra de sessão anterior (ou de uma busca manual feita na mesma aba): um
    # chip de shortHash já aplicado filtraria a grade junto com o do 1º produto.
    residuais = _limpar_chips_shorthash(driver)
    if residuais:
        logger.info(f"  Filtro(s) de shortHash residual(is) removido(s) antes de começar: {residuais}")

    total_atual, _ = _ler_total_produtos_categoria(driver)
    if total_atual is None:
        total_atual = _contar_paginando(driver, SELETOR_LINHAS_CATEGORIA, SELETOR_PROXIMA_PAGINA_CATEGORIA)
        if total_atual is not None:
            logger.warning("  ⚠ Contador de paginação indisponível — contagem inicial via páginas (mais lento)")
        else:
            logger.warning(
                "  ⚠ Não foi possível determinar o total de produtos na coleção — "
                "cliques serão aceitos sem verificação de contagem"
            )
    logger.info(f"  Total de produtos na coleção antes de adicionar: {total_atual}")

    adicionados = []
    nao_encontrados = []
    sem_efeito = []
    codigo_nao_bate = []
    abortado = False

    # Encadeia a linha aceita de uma chamada como `referencia_anterior` da
    # próxima — só tem efeito quando LIMPAR_FILTRO_ENTRE_BUSCAS=False (ver
    # docstring de `_tentar_adicionar_um`); com a flag True, cada chamada já
    # limpa e confirma a limpeza antes de retornar, então `linha_anterior`
    # não influencia em nada (a tela já está "zerada" na próxima chamada).
    linha_anterior = None

    for short_hash in codigo_por_hash:
        codigo_esperado = codigo_por_hash[short_hash]
        status, total_atual, linha_anterior = _tentar_adicionar_um(
            driver, short_hash, codigo_esperado, seletores, total_atual, linha_anterior
        )

        if status == 'campo_ausente':
            abortado = True
            break
        elif status == 'adicionado':
            adicionados.append(short_hash)
        elif status == 'sem_efeito':
            sem_efeito.append(short_hash)
        elif status == 'codigo_nao_bate':
            codigo_nao_bate.append(short_hash)
        else:
            nao_encontrados.append(short_hash)

    # Retry único para quem teve clique sem efeito (não para "não encontrado"
    # nem "código não bate", que já significam que repetir a MESMA busca não
    # mudaria o resultado).
    if sem_efeito and not abortado:
        logger.info(f"Retentando {len(sem_efeito)} produto(s) sem efeito na primeira tentativa...")
        pendentes, sem_efeito = sem_efeito, []
        for short_hash in pendentes:
            codigo_esperado = codigo_por_hash[short_hash]
            status, total_atual, linha_anterior = _tentar_adicionar_um(
                driver, short_hash, codigo_esperado, seletores, total_atual, linha_anterior
            )
            if status == 'adicionado':
                adicionados.append(short_hash)
            elif status == 'codigo_nao_bate':
                codigo_nao_bate.append(short_hash)
            elif status == 'campo_ausente':
                sem_efeito.append(short_hash)
                break
            else:
                sem_efeito.append(short_hash)

    logger.info(
        f"✓ Adição concluída — {len(adicionados)} adicionado(s), {len(nao_encontrados)} não encontrado(s), "
        f"{len(sem_efeito)} sem efeito, {len(codigo_nao_bate)} com código divergente"
    )
    if nao_encontrados:
        logger.warning(f"  shortHashes não encontrados: {nao_encontrados}")
    if sem_efeito:
        logger.warning(f"  shortHashes sem efeito mesmo após retry: {sem_efeito}")
    if codigo_nao_bate:
        logger.warning(f"  shortHashes com código divergente do SKU (não adicionados): {codigo_nao_bate}")

    return {
        "adicionados": adicionados,
        "nao_encontrados": nao_encontrados,
        "sem_efeito": sem_efeito,
        "codigo_nao_bate": codigo_nao_bate,
    }


def conferir_resultado_final(driver, lista_shorthashes):
    """Recarrega a página da coleção e confere se o total de produtos bate
    com o esperado (len(lista_shorthashes)). Se divergir, tenta (best-effort)
    identificar quais shortHashes da planilha não aparecem na tabela
    'Produtos da categoria' — ver ressalva em `_listar_shorthashes_faltando`
    sobre quando essa parte é pulada por não ser confiável."""
    logger.info("Conferindo resultado final da coleção...")
    driver.refresh()

    pagina_ok = _tentar_localizar(driver, [_H6_CATEGORIA], timeout=TIMEOUT_ELEMENTO)
    if pagina_ok is None:
        logger.error("  ✗ Não foi possível recarregar a página da coleção para conferência final")
        return

    esperado = len(lista_shorthashes)

    total_real, _ = _ler_total_produtos_categoria(driver)
    if total_real is None:
        total_real = _contar_paginando(driver, SELETOR_LINHAS_CATEGORIA, SELETOR_PROXIMA_PAGINA_CATEGORIA)

    if total_real is None:
        logger.error("  ✗ Não foi possível determinar o total real de produtos na coleção para conferência final")
        return

    diferenca = esperado - total_real
    if diferenca != 0:
        logger.error(
            f"  ✗ DIVERGÊNCIA FINAL — esperado: {esperado} | encontrado na coleção: {total_real} | "
            f"diferença: {diferenca}"
        )
        faltando = _listar_shorthashes_faltando(
            driver, lista_shorthashes, SELETOR_LINHAS_CATEGORIA, SELETOR_PROXIMA_PAGINA_CATEGORIA
        )
        if faltando:
            logger.error(f"  shortHashes da planilha não localizados na coleção: {faltando}")
    else:
        logger.info(f"  ✓ Contagem final confere: {total_real} produto(s) na coleção")


def salvar(driver):
    """Clica no botão 'Salvar' e valida sucesso."""
    logger.info("Salvando alterações...")

    clicou = _tentar_clicar(
        driver,
        ["//button[normalize-space(.)='Salvar']", "//*[normalize-space(text())='Salvar']"],
        descricao="botão Salvar",
    )
    if not clicou:
        raise RuntimeError("Não foi possível clicar no botão 'Salvar'")

    sucesso = _tentar_localizar(
        driver,
        [
            "//*[contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZÇÃÕ','abcdefghijklmnopqrstuvwxyzçãõ'),'sucesso')]",
            "//*[contains(normalize-space(.),'salvo') or contains(normalize-space(.),'salva')]",
        ],
        timeout=TIMEOUT_ELEMENTO,
    )
    if sucesso:
        logger.info("✓ Alterações salvas com sucesso")
    else:
        logger.warning("  ⚠ Não foi possível confirmar mensagem de sucesso após salvar")


def configurar_chrome():
    opcoes = Options()
    # 'eager' retorna o controle assim que o DOM está pronto/interativo, sem
    # esperar recursos secundários (imagens, etc.) terminarem de carregar —
    # todas as esperas do bot já são explícitas (via _tentar_localizar), não
    # dependem do evento de "página totalmente carregada".
    opcoes.page_load_strategy = 'eager'
    opcoes.add_argument("--start-maximized")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_argument("--log-level=3")
    # A tabela de resultados carrega miniaturas de produto a cada busca — como
    # o bot nunca olha pra imagem nenhuma, desabilitar o carregamento poupa
    # banda/tempo de render sem afetar nenhuma checagem do fluxo.
    opcoes.add_experimental_option(
        "prefs", {"profile.managed_default_content_settings.images": 2}
    )
    return opcoes


def main():
    logger.info("=" * 60)
    logger.info("INICIANDO ROBÔ DE GERENCIAMENTO DE COLEÇÕES - SPORTBAY")
    logger.info("=" * 60)

    if not USUARIO or not SENHA:
        logger.critical(
            "Credenciais não encontradas. Defina SPORTBAY_USUARIO e SPORTBAY_SENHA "
            "em um arquivo .env na pasta do script."
        )
        return

    try:
        logger.info(f"Lendo planilha: {NOME_PLANILHA}")
        df = pd.read_excel(NOME_PLANILHA, header=None, skiprows=1)
    except FileNotFoundError:
        logger.critical(f"Planilha não encontrada: {NOME_PLANILHA}")
        return
    except Exception as e:
        logger.critical(f"Erro ao ler planilha: {e}")
        return

    if df.empty:
        logger.critical("Planilha está vazia")
        return

    nome_colecao = str(df.iloc[0, COL_COLECAO_IDX]).strip()
    produtos_bruto = [
        (str(short_hash).strip(), str(codigo).strip())
        for short_hash, codigo in zip(df.iloc[:, COL_SHORTHASH_IDX], df.iloc[:, COL_CODIGO_IDX])
        if str(short_hash).strip() and str(short_hash).strip().lower() != 'nan'
    ]

    # Dedupe por shortHash preservando ordem — shortHash duplicado na planilha
    # faria o bot tentar adicionar (ou contar) o mesmo produto duas vezes.
    vistos = set()
    lista_produtos = []
    duplicados = []
    for short_hash, codigo in produtos_bruto:
        if short_hash in vistos:
            duplicados.append(short_hash)
        else:
            vistos.add(short_hash)
            lista_produtos.append((short_hash, codigo))
    if duplicados:
        logger.warning(
            f"⚠ {len(duplicados)} shortHash(es) duplicado(s) na planilha, descartados: {duplicados}"
        )

    lista_shorthashes = [h for h, _ in lista_produtos]

    logger.info(f"Coleção alvo: '{nome_colecao}'")
    logger.info(f"Total de shortHashes a adicionar: {len(lista_shorthashes)}")

    driver = None
    try:
        # Selenium Manager (embutido no Selenium 4.6+) resolve e baixa o
        # chromedriver certo automaticamente — sem precisar do
        # ChromeDriverManager().install() (webdriver-manager), que fazia uma
        # checagem de versão própria a cada início.
        driver = webdriver.Chrome(options=configurar_chrome())

        login(driver)
        navegar_para_lista_categorias(driver)
        abrir_colecao(driver, nome_colecao)
        remover_todos_produtos_atuais(driver)
        limpar_filtro_merchant(driver)
        adicionar_produtos(driver, lista_produtos)
        salvar(driver)
        conferir_resultado_final(driver, lista_shorthashes)

        logger.info("=" * 60)
        logger.info("✓ FLUXO CONCLUÍDO")
        logger.info("=" * 60)

    except Exception as e:
        logger.critical(f"Erro durante execução do fluxo: {e}")

    finally:
        if driver:
            driver.quit()
            logger.info("Chrome fechado")
        _logar_resumo_tempos()


if __name__ == "__main__":
    main()
