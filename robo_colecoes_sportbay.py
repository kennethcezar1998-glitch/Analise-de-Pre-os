"""
Bot de automação — Gerenciamento de Produtos em Coleções (Sportbay Marketplace)

Remove todos os produtos atualmente associados a uma coleção e adiciona os
produtos (por shortHash) listados na planilha de resultado, excluindo o
merchant "Sportbay" da busca de associação manual.
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
NOME_PLANILHA = os.path.join(PASTA_SCRIPT, "Resultado_Menor_Preco_Sportbay.xlsx")

COL_CODIGO_IDX = 0      # Coluna A (Código Fabrica)
COL_SHORTHASH_IDX = 7   # Coluna H
COL_COLECAO_IDX = 9     # Coluna J

TIMEOUT_ELEMENTO = 15
TIMEOUT_LOGIN = 25
MAX_ITERACOES_REMOCAO = 1000

# Se True (comportamento atual, default): depois de cada produto, limpa o
# campo de shortHash e dá Enter antes do próximo — dispara uma SEGUNDA busca
# completa ao backend, só pra "resetar" a tela.
# Se False: pula essa busca de limpeza e vai direto pro próximo shortHash. A
# trava de staleness em `_aguardar_resultado_novo` já garante a mesma
# segurança (só aceita resultado novo depois que o anterior sai do DOM), já
# que a referência do produto processado é passada explicitamente adiante
# como `referencia_anterior` — ver `_tentar_adicionar_um` e `adicionar_produtos`.
LIMPAR_FILTRO_ENTRE_BUSCAS = True

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


def _tratar_modal_confirmacao(driver, timeout=1):
    """Se aparecer um modal de confirmação (ex.: 'Deseja remover?'), clica em confirmar.

    Checagem IMEDIATA via `find_elements` (sem WebDriverWait) — o caso comum
    é NÃO existir modal nenhum (a remoção é instantânea ao clicar no ícone),
    e essa função é chamada a cada produto removido. Um `WebDriverWait(0.5)`
    aqui custava 0.5s por produto só pra confirmar a ausência do modal; um
    `find_elements` puro retorna na hora, sem esperar nada.

    `timeout` foi mantido na assinatura por compatibilidade com quem chama,
    mas não é mais usado (não há espera nesta versão).
    """
    seletor_confirmar = (
        "//div[contains(@class,'modal')]//button["
        "contains(normalize-space(.),'Sim') or contains(normalize-space(.),'Confirmar') or contains(normalize-space(.),'OK')]"
        " | //*[contains(@class,'modal') and contains(@class,'show')]//button[contains(normalize-space(.),'Sim')]"
    )
    elementos = driver.find_elements(By.XPATH, seletor_confirmar)
    for elemento in elementos:
        try:
            if not (elemento.is_displayed() and elemento.is_enabled()):
                continue
            elemento.click()
            logger.debug("  → Modal de confirmação tratado")
            time.sleep(0.5)
            return True
        except Exception:
            continue
    return False


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


def remover_todos_produtos_atuais(driver):
    """Remove, um a um, todos os produtos já associados à coleção, percorrendo
    todas as páginas da listagem 'Produtos da categoria'."""
    logger.info("Removendo produtos atuais da coleção...")

    # Usa //h6[...] (não //*[...]) de propósito: com //*, o texto "Produtos da
    # categoria" também "existe" em todo ancestral do heading (div pai, avô,
    # até o body, já que `.` pega o texto de toda a subárvore) — isso faz
    # `following::table[1]` virar ambíguo (cada ancestral resolve para um
    # "próximo table" possivelmente diferente). Travando no <h6> específico,
    # há só um nó de contexto e o `following::` sempre aponta pra tabela certa.
    secao_produtos = _tentar_localizar(
        driver,
        ["//h6[contains(normalize-space(.),'Produtos da categoria')]"],
        timeout=TIMEOUT_ELEMENTO,
    )
    if secao_produtos is None:
        raise RuntimeError("Seção 'Produtos da categoria' não encontrada na página de edição")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", secao_produtos)

    seletores_remover_linha = [
        # Ícone real observado no DOM: <svg data-testid="RemoveCircleIcon">
        "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::table[1]//*[@data-testid='RemoveCircleIcon']",
        "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::table[1]//button[contains(@title,'Remov') or contains(@aria-label,'Remov')]",
    ]
    seletor_linhas_categoria = "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::table[1]/tbody/tr"
    seletor_proxima_pagina = (
        "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::*"
        "[self::a or self::button][contains(@aria-label,'próxima') or contains(@aria-label,'Próxima')"
        " or contains(@class,'next')][not(contains(@class,'disabled'))]"
    )

    total_removidos = 0
    for _ in range(MAX_ITERACOES_REMOCAO):
        with _medir('remocao_item'):
            botao_remover = _tentar_localizar(driver, seletores_remover_linha, timeout=4)

            if botao_remover is None:
                proxima = _tentar_localizar(driver, [seletor_proxima_pagina], timeout=3, clicavel=True)
                if proxima:
                    try:
                        _clicar_proxima_pagina_e_aguardar(driver, proxima, seletor_linhas_categoria, timeout_fallback=1)
                        continue
                    except Exception:
                        break
                break

            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_remover)
                botao_remover.click()
            except StaleElementReferenceException:
                # Relocaliza e tenta de novo, em vez de só `continue` — um `continue`
                # aqui reinicia o loop externo silenciosamente, então uma falha
                # repetida de stale reference passaria despercebida como se nada
                # tivesse acontecido, arriscando pular o produto sem log nenhum.
                botao_remover = _tentar_localizar(driver, seletores_remover_linha, timeout=4)
                if botao_remover is None:
                    break
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_remover)
                    botao_remover.click()
                except Exception as e:
                    logger.warning(f"  ⚠ Falha ao remover produto após relocalizar (stale): {e}")
                    continue
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", botao_remover)

            # Checagem rápida de modal de confirmação — na maioria dos casos não
            # existe modal (remoção é instantânea ao clicar no ícone), então o
            # timeout aqui é propositalmente curto para não acumular espera à toa
            # a cada produto removido.
            _tratar_modal_confirmacao(driver, timeout=0.5)

            try:
                WebDriverWait(driver, 3, poll_frequency=0.1).until(EC.staleness_of(botao_remover))
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


def _aguardar_resultado_novo(driver, seletor_linhas, referencia_antiga, timeout=10):
    """Aguarda a tabela de resultados mostrar um resultado NOVO para a busca
    atual, distinguindo-o do resultado da busca ANTERIOR via staleness — não
    via texto: o shortHash não aparece em nenhuma coluna visível da tabela
    (Nome/SKU/Situação/Merchant/Criação/Estado), então comparar texto nunca
    bate e sempre estoura timeout (foi tentado e confirmado nesta planilha).

    Bug que isso corrige: se o filtro anterior ainda não foi limpo no
    servidor no momento em que o novo shortHash é buscado, a tabela pode
    continuar mostrando, por um instante, a única linha do PRODUTO ANTERIOR.
    Uma checagem de "1 linha" sem mais nada (versão antiga,
    `_aguardar_resultado_unico`) aceitava essa linha errada, e o bot clicava
    em adicionar o produto repetido/errado — daí o log dizer "64 adicionados"
    com só 51 produtos reais na coleção.

    Se `referencia_antiga` existia (era a única linha antes de disparar esta
    busca), só aceita depois que ELA ficar stale — prova de que o DOM
    realmente re-renderizou por causa desta busca, não da anterior. Depois
    disso (ou se não havia `referencia_antiga`), basta a tabela ter 1 linha.

    Retorna (ok, contagem, linha_nova) — linha_nova é o WebElement da linha
    aceita (útil para o chamador rastrear staleness dela depois), ou None.

    Cada iteração do polling faz UMA chamada via `_consultar_estado_resultado`
    (1 round-trip ao chromedriver) em vez de 2-3 chamadas separadas
    (find_elements + execute_script pra staleness/conteúdo) — essa função é
    chamada até dezenas de vezes por shortHash, então cada round-trip a menos
    por iteração conta.
    """
    fim = time.time() + timeout

    if referencia_antiga is not None:
        ref_stale, _, _, _ = _consultar_estado_resultado(driver, seletor_linhas, referencia_antiga)
        if not ref_stale:
            ficou_stale = False
            while time.time() < fim:
                ref_stale, _, _, _ = _consultar_estado_resultado(driver, seletor_linhas, referencia_antiga)
                if ref_stale:
                    ficou_stale = True
                    break
                time.sleep(0.1)
            if not ficou_stale:
                _, contagem, _, _ = _consultar_estado_resultado(driver, seletor_linhas)
                return False, contagem, None

    ultima_contagem = None
    while time.time() < fim:
        # Além de contagem==1, exige que a linha já tenha conteúdo renderizado
        # — o <tr> pode aparecer no DOM (contando como "1 linha") um instante
        # antes das células serem preenchidas (skeleton/loading), e aceitar
        # essa linha vazia cedo demais fazia a validação de SKU/Merchant
        # (feita logo depois, em _tentar_adicionar_um) ler '' e '' e falhar
        # com "código não bate" mesmo quando o produto era o certo.
        _, ultima_contagem, tem_conteudo, linha = _consultar_estado_resultado(driver, seletor_linhas)
        if ultima_contagem == 1 and tem_conteudo:
            return True, ultima_contagem, linha
        time.sleep(0.1)
    return False, ultima_contagem, None


_JS_CONSULTAR_ESTADO_RESULTADO = """
    const seletor = arguments[0];
    const refAntiga = arguments[1];
    const resultado = document.evaluate(
        seletor, document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null
    );
    const contagem = resultado.snapshotLength;
    let refStale = true;
    if (refAntiga) {
        try { refStale = !refAntiga.isConnected; } catch (e) { refStale = true; }
    }
    let temConteudo = false;
    let linha = null;
    if (contagem === 1) {
        linha = resultado.snapshotItem(0);
        temConteudo = (linha.textContent || '').trim() !== '';
    }
    return [refStale, contagem, temConteudo, temConteudo ? linha : null];
"""


def _consultar_estado_resultado(driver, seletor_linhas, referencia_antiga=None):
    """Uma única chamada ao navegador (via `document.evaluate`) que retorna,
    de uma vez: se `referencia_antiga` já saiu do DOM (`.isConnected`), a
    contagem atual de linhas, se a linha (quando há exatamente 1) já tem
    conteúdo renderizado, e a própria linha (ou None). Usado dentro do laço
    de polling de `_aguardar_resultado_novo` para colapsar em 1 round-trip o
    que antes eram várias chamadas separadas por iteração.
    """
    try:
        ref_stale, contagem, tem_conteudo, linha = driver.execute_script(
            _JS_CONSULTAR_ESTADO_RESULTADO, seletor_linhas, referencia_antiga
        )
    except StaleElementReferenceException:
        # referencia_antiga não pôde nem ser resolvida como argumento do
        # script — equivale a já estar stale; refaz a consulta sem ela.
        linhas = driver.find_elements(By.XPATH, seletor_linhas)
        contagem = len(linhas)
        tem_conteudo = False
        linha = None
        if contagem == 1:
            try:
                texto = driver.execute_script("return arguments[0].textContent;", linhas[0]) or ""
                tem_conteudo = texto.strip() != ""
                linha = linhas[0] if tem_conteudo else None
            except StaleElementReferenceException:
                pass
        ref_stale = True
    return ref_stale, contagem, tem_conteudo, linha


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


def _extrair_sku_e_merchant(driver, linha):
    """Lê o texto das células de SKU (3ª coluna) e Merchant (5ª coluna) da
    linha de resultado, via textContent (não .text — a linha pode estar fora
    da área visível do scroll do simplebar, e .text retorna vazio nesse caso).
    Colunas da tabela, em ordem: [checkbox, Nome, SKU, Situação, Merchant,
    Criação, Estado, ação] — confirmado no DOM real da 'Seleção Manual de
    Produtos'.

    Um único execute_script busca as duas células (via getElementsByTagName,
    índices 2 e 4 = 3ª e 5ª coluna) e devolve os dois textContent de uma vez
    — 1 round-trip ao chromedriver, em vez de 2 find_element + 2 execute_script."""
    try:
        sku, merchant = driver.execute_script(
            """
            const linha = arguments[0];
            const tds = linha.getElementsByTagName('td');
            return [
                tds[2] ? tds[2].textContent : '',
                tds[4] ? tds[4].textContent : '',
            ];
            """,
            linha,
        )
    except (StaleElementReferenceException, Exception):
        return "", ""
    return (sku or "").strip(), (merchant or "").strip()


def _codigo_bate_com_sku(codigo_planilha, sku_site, merchant_site):
    """Compara o código da Coluna A da planilha com o SKU mostrado na linha.

    Normalizações:
      - O prefixo 'PAI-' do SKU do site é descartado antes de comparar (a
        planilha nunca traz esse prefixo).
      - Se o código da planilha terminar com 'A' e o SKU do site (já sem o
        prefixo) bater com o código SEM o 'A', só é considerado válido se o
        Merchant da linha for 'Sportbay' — fora esse caso o sufixo importa.
    """
    codigo = str(codigo_planilha).strip()
    if codigo.endswith('.0'):  # artefato de leitura como float (ex.: pandas)
        codigo = codigo[:-2]
    if not codigo or codigo.lower() == 'nan':
        return True  # sem código pra validar, não bloqueia

    sku = sku_site.strip().upper()
    if sku.startswith('PAI-'):
        sku = sku[len('PAI-'):]

    codigo_up = codigo.upper()
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

    seletor_paginacao = (
        "//h6[contains(normalize-space(.),'Produtos da categoria')]"
        "/following::div[contains(@class,'MuiTablePagination-root')][1]"
    )
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
    seletores_botao_adicionar = seletores["botao_adicionar"]

    # Busca o shortHash — com 1 retentativa se a tabela não convergir para 1
    # resultado novo na primeira tentativa (ex.: Enter não disparou a busca a
    # tempo, ou a tabela ficou mostrando o estado anterior). Repetir a mesma
    # busca do zero (limpar/redigitar/Enter de novo) resolve a maioria desses
    # casos sem precisar desistir do shortHash de cara.
    resultado_ok, contagem, linha_nova = False, None, None
    with _medir('busca'):
        for tentativa_busca in (1, 2):
            # Referência do que está na tela ANTES de disparar esta busca — usada
            # por `_aguardar_resultado_novo` para confirmar via staleness que o
            # resultado seguinte é realmente novo (não a linha do produto
            # anterior, ainda não descartada pelo servidor). Na 1ª tentativa,
            # se a limpeza entre buscas está desligada, usa a referência já
            # conhecida (evita uma consulta ao DOM e é o ponto correto de
            # comparação, já que a tela não foi limpa desde a busca anterior).
            # Na 2ª tentativa (retry) sempre reconsulta — o estado pode ter
            # mudado depois da 1ª tentativa falhar.
            if tentativa_busca == 1 and not LIMPAR_FILTRO_ENTRE_BUSCAS and referencia_anterior is not None:
                referencia_antiga = referencia_anterior
            else:
                referencia_antiga = _linha_unica_atual(driver, seletor_linhas_resultado)

            campo = _tentar_localizar(driver, seletores_campo_filtro, timeout=TIMEOUT_ELEMENTO)
            if campo is None:
                logger.error("  ✗ Campo 'Filtrar por shortHash' não encontrado — abortando adição")
                return 'campo_ausente', total_atual, None

            _preencher_campo(driver, campo, short_hash)
            campo.send_keys(Keys.ENTER)  # a busca é disparada ao apertar Enter, não ao digitar

            resultado_ok, contagem, linha_nova = _aguardar_resultado_novo(
                driver, seletor_linhas_resultado, referencia_antiga, timeout=10
            )
            if resultado_ok:
                break

            if tentativa_busca == 1:
                logger.warning(
                    f"  ⚠ shortHash '{short_hash}': tabela não convergiu para 1 resultado novo "
                    f"(contagem no timeout: {contagem}) — pesquisando novamente..."
                )

    status = None
    novo_total = total_atual

    if not resultado_ok:
        logger.warning(
            f"  ⚠ shortHash '{short_hash}': tabela não convergiu para 1 resultado novo mesmo após "
            f"repetir a busca (contagem no timeout: {contagem}) — pulando"
        )
        status = 'nao_encontrado'
    else:
        # Trava extra ANTES de clicar: confere se o SKU da linha realmente
        # corresponde ao código da planilha (Coluna A), tolerando o prefixo
        # "PAI-" e, só para Merchant "Sportbay", a ausência do sufixo "A".
        with _medir('validacao_sku'):
            sku_site, merchant_site = (
                _extrair_sku_e_merchant(driver, linha_nova) if linha_nova is not None else ("", "")
            )
            codigo_bate = _codigo_bate_com_sku(codigo_esperado, sku_site, merchant_site)

        if not codigo_bate:
            logger.warning(
                f"  ⚠ shortHash '{short_hash}': código '{codigo_esperado}' não bate com o SKU da linha "
                f"('{sku_site}', Merchant: '{merchant_site}') — pulando para não arriscar adicionar o produto errado"
            )
            status = 'codigo_nao_bate'
        else:
            with _medir('clique'):
                botao_adicionar = _tentar_localizar(driver, seletores_botao_adicionar, timeout=4, clicavel=True)
                if botao_adicionar is None:
                    logger.warning(f"  ⚠ shortHash não encontrado na busca: {short_hash}")
                    status = 'nao_encontrado'
                else:
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao_adicionar)
                        botao_adicionar.click()
                    except StaleElementReferenceException:
                        # Relocaliza e tenta de novo, em vez de desistir direto — o
                        # ícone pode ter sido re-renderizado entre localizar e clicar.
                        botao_adicionar = _tentar_localizar(driver, seletores_botao_adicionar, timeout=4, clicavel=True)
                        if botao_adicionar is None:
                            status = 'nao_encontrado'
                        else:
                            try:
                                botao_adicionar.click()
                            except Exception as e:
                                logger.warning(f"  ⚠ Falha ao clicar em adicionar para {short_hash} (após relocalizar): {e}")
                                status = 'nao_encontrado'
                    except ElementClickInterceptedException:
                        driver.execute_script("arguments[0].click();", botao_adicionar)
                    except Exception as e:
                        logger.warning(f"  ⚠ Falha ao clicar em adicionar para {short_hash}: {e}")
                        status = 'nao_encontrado'

            if status is None:  # clique aconteceu sem exceção — verificar efeito real
                with _medir('confirmacao_contador'):
                    if total_atual is None:
                        logger.info(f"  ➕ Produto adicionado: {short_hash} (sem verificação de contagem)")
                        status = 'adicionado'
                    else:
                        fim = time.time() + 8
                        confirmado = False
                        cache_paginacao = None
                        while time.time() < fim:
                            lido, cache_paginacao = _ler_total_produtos_categoria(driver, cache_paginacao)
                            if lido is not None and lido == total_atual + 1:
                                novo_total = lido
                                confirmado = True
                                break
                            time.sleep(0.15)
                        if confirmado:
                            logger.info(f"  ➕ Produto adicionado: {short_hash} (total {total_atual} → {novo_total})")
                            status = 'adicionado'
                        else:
                            logger.warning(
                                f"  ⚠ shortHash '{short_hash}': clique sem efeito "
                                f"(total de produtos continuou em {total_atual})"
                            )
                            status = 'sem_efeito'

    # Limpa o campo de filtro antes da próxima busca — só depois de tudo acima
    # (confirmação ou desistência), nunca em paralelo com uma requisição de
    # adição ainda em andamento no backend. Com LIMPAR_FILTRO_ENTRE_BUSCAS=False,
    # pula essa busca de limpeza inteira: digitar o próximo shortHash já
    # substitui o valor do campo (_preencher_campo define o valor direto, não
    # concatena), e a trava de staleness em `_aguardar_resultado_novo` (usando
    # `linha_nova` como `referencia_anterior` na próxima chamada) continua
    # garantindo que o resultado antigo não seja aceito como se fosse novo.
    if LIMPAR_FILTRO_ENTRE_BUSCAS:
        with _medir('limpeza_filtro'):
            try:
                campo_atual = _tentar_localizar(driver, seletores_campo_filtro, timeout=3)
                if campo_atual:
                    _preencher_campo(driver, campo_atual, "")
                    campo_atual.send_keys(Keys.ENTER)
                    if not _aguardar_linha_sumir(linha_nova, timeout=10):
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

    `lista_produtos` é uma lista de tuplas (short_hash, codigo_fabrica), onde
    codigo_fabrica vem da Coluna A da planilha e é usado para validar o SKU
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
    seletores = {
        "linhas_resultado": f"{seletor_tabela}/tbody/tr",
        "campo_filtro": ["//input[contains(@placeholder,'shortHash')]"],
        "botao_adicionar": [
            # Ícone real observado no DOM: <svg data-testid="AddBoxIcon" aria-label="Adicionar produto">
            f"{seletor_tabela}//*[@data-testid='AddBoxIcon']",
            f"{seletor_tabela}//*[@aria-label='Adicionar produto']",
        ],
    }

    total_atual, _ = _ler_total_produtos_categoria(driver)
    if total_atual is None:
        seletor_linhas_categoria = "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::table[1]/tbody/tr"
        seletor_proxima_categoria = (
            "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::*"
            "[self::a or self::button][contains(@aria-label,'próxima') or contains(@aria-label,'Próxima')"
            " or contains(@class,'next')][not(contains(@class,'disabled'))]"
        )
        total_atual = _contar_paginando(driver, seletor_linhas_categoria, seletor_proxima_categoria)
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

    pagina_ok = _tentar_localizar(
        driver,
        ["//h6[contains(normalize-space(.),'Produtos da categoria')]"],
        timeout=TIMEOUT_ELEMENTO,
    )
    if pagina_ok is None:
        logger.error("  ✗ Não foi possível recarregar a página da coleção para conferência final")
        return

    esperado = len(lista_shorthashes)
    seletor_linhas_categoria = "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::table[1]/tbody/tr"
    seletor_proxima_categoria = (
        "//h6[contains(normalize-space(.),'Produtos da categoria')]/following::*"
        "[self::a or self::button][contains(@aria-label,'próxima') or contains(@aria-label,'Próxima')"
        " or contains(@class,'next')][not(contains(@class,'disabled'))]"
    )

    total_real, _ = _ler_total_produtos_categoria(driver)
    if total_real is None:
        total_real = _contar_paginando(driver, seletor_linhas_categoria, seletor_proxima_categoria)

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
            driver, lista_shorthashes, seletor_linhas_categoria, seletor_proxima_categoria
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
