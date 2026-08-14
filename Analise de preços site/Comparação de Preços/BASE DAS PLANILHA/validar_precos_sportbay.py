"""
Robô de VALIDAÇÃO de preços — confere se os preços da planilha
"Base dos produtos para alteração.xlsx" foram realmente refletidos no
site da Sportbay (www.sportbay.com.br).

Fluxo por produto (coluna A da planilha):
  1. Lê o código na coluna A. Se terminar com sufixo "A", remove o sufixo
     antes de pesquisar (ex: "116086A" -> pesquisa "116086").
  2. Abre a Sportbay e pesquisa o código na barra de busca do site.
  3. Na grade de resultados, clica no primeiro produto. Se houver mais de
     um resultado, prioriza o anúncio vendido "por Sportbay" (1ª parte),
     em vez de anúncios de marketplace de terceiros.
  4. Na página do produto, extrai o preço "no Pix".
  5. Compara com o preço esperado (coluna "Sugestão de preço final").
     Pequenas diferenças de centavos (arredondamento) são toleradas.
  6. Se o preço da variação padrão (cor/tamanho selecionados por padrão)
     ficar DIVERGENTE, o robô clica nas demais variações de cor e tamanho
     do produto e reconfere o preço em cada uma, até achar uma que bata
     com o esperado. Só mantém o status DIVERGENTE se nenhuma variação
     tiver o preço correto.

Resultado é salvo em "Resultado_Validacao_Precos.xlsx" na mesma pasta.
"""

import logging
import os
import re
import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# Diretório do próprio script, para que os caminhos funcionem
# independentemente de onde o script for executado (ex: quando chamado
# por outro processo/bot a partir de uma pasta diferente).
DIR_SCRIPT = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURAÇÕES DE LOGGING ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(DIR_SCRIPT, 'validar_precos_sportbay.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- CONFIGURAÇÕES ---
ARQUIVO_ENTRADA = os.path.join(DIR_SCRIPT, 'Base dos produtos para alteração.xlsx')
ARQUIVO_SAIDA = os.path.join(DIR_SCRIPT, 'Resultado_Validacao_Precos.xlsx')

COL_CODIGO = 0          # coluna A: "PAI" (código do produto)
COL_PRODUTO = 1          # coluna B: "PRODUTO"
COL_PRECO_ESPERADO = 4   # coluna E: "Sugestão de preço final"

TOLERANCIA_PRECO = 0.05  # diferença (em R$) tolerada por arredondamento do site
TIMEOUT_ELEMENTO = 10
TEMPO_ESPERA_ENTRE_PRODUTOS = 2
MAX_TENTATIVAS = 2


# --- FUNÇÕES AUXILIARES (preço / código) ---

def extrair_menor_preco(texto):
    """Extrai o menor preço em formato R$ do texto (aceita formato BR e fallback EN)."""
    try:
        matches = re.findall(r'R\$\s*([\d\.,]+)', texto)
        menor_valor = float('inf')
        for match in matches:
            try:
                if ',' in match:
                    valor_str = match.replace('.', '').replace(',', '.')
                else:
                    valor_str = match
                valor = float(valor_str)
                if valor > 0 and valor < menor_valor:
                    menor_valor = valor
            except ValueError:
                continue
        return menor_valor if menor_valor != float('inf') else None
    except Exception as e:
        logger.warning(f"Erro ao extrair preço de '{texto}': {e}")
        return None


def codigo_termina_com_a(codigo):
    """Verifica se o código termina com o sufixo 'A' (case-insensitive)."""
    codigo = (codigo or "").strip()
    return len(codigo) > 1 and codigo[-1].upper() == 'A'


def configurar_chrome():
    """Configura opções do Chrome com máxima compatibilidade."""
    opcoes = Options()
    opcoes.add_argument("--start-maximized")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")
    opcoes.add_argument("--disable-blink-features=AutomationControlled")
    opcoes.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    opcoes.add_argument("--log-level=3")
    return opcoes


# --- BUSCA E SELEÇÃO DE PRODUTO ---

def buscar_produto(driver, codigo_produto):
    """Pesquisa o código na barra de busca real do site (com fallback via URL direta)."""
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    try:
        driver.get("https://www.sportbay.com.br/")

        campo_busca = wait.until(
            EC.element_to_be_clickable((By.ID, "search_user_component"))
        )
        campo_busca.clear()
        campo_busca.send_keys(codigo_produto)

        try:
            botao_busca = wait.until(EC.element_to_be_clickable((By.ID, "seach_button")))
            botao_busca.click()
        except Exception:
            botao_busca = driver.find_element(By.ID, "seach_button")
            driver.execute_script("arguments[0].click();", botao_busca)

        try:
            wait.until(EC.url_contains("busca"))
        except Exception:
            campo_busca.send_keys(Keys.RETURN)
            wait.until(EC.url_contains("busca"))

        wait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '/p')]"))
        )
        return True
    except Exception as e:
        logger.debug(f"  Busca via barra de pesquisa falhou ({e}), tentando via URL direta...")
        try:
            url_busca = f"https://www.sportbay.com.br/busca?q={codigo_produto}"
            driver.get(url_busca)
            wait.until(
                EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, '/p')]"))
            )
            return True
        except Exception:
            return False


def selecionar_link_produto(driver):
    """
    Escolhe qual produto da grade de resultados abrir.

    Regra: percorre os resultados na ordem em que aparecem na grade e
    prioriza o primeiro anúncio vendido "por Sportbay" (1ª parte) —
    identificado pelo <span aria-label="Sportbay"> dentro do card, como
    visto no DevTools. Se nenhum resultado for vendido pela Sportbay
    (ex: só há anúncios de marketplace), cai para o primeiro produto da
    grade, que é o comportamento padrão.

    Retorna a URL do produto escolhido, ou None se não houver resultados.
    """
    elementos = driver.find_elements(By.XPATH, "//a[contains(@href, '/p')]")

    candidatos = []
    vistos = set()
    for el in elementos:
        href = el.get_attribute('href') or ''
        if not (href.endswith('/p') or '/p?' in href):
            continue
        href_base = href.split('?')[0]
        if href_base in vistos:
            continue
        vistos.add(href_base)

        vendido_por_sportbay = False
        try:
            el.find_element(By.XPATH, ".//span[@aria-label='Sportbay']")
            vendido_por_sportbay = True
        except NoSuchElementException:
            pass

        candidatos.append((href, vendido_por_sportbay))

    if not candidatos:
        return None

    for href, eh_sportbay in candidatos:
        if eh_sportbay:
            return href

    # Nenhum resultado marcado como "por Sportbay" -> usa o primeiro da grade
    return candidatos[0][0]


def extrair_preco_pix_na_pagina(driver):
    """
    Extrai o preço "no Pix" na página do produto.

    Estrutura real (DevTools):
        <span class="price-main">R$&nbsp;144,90</span>
        <span class="priceObs">no <strong>Pix</strong></span>
    dentro do mesmo <p>. Usamos essa relação para não pegar preço de parcela.
    """
    seletores_pix = [
        "//strong[contains(translate(text(),'PIX','pix'),'pix')]/ancestor::p[1]//span[contains(@class,'price-main')]",
        "//*[contains(@class,'priceObs')][.//strong[contains(translate(text(),'PIX','pix'),'pix')] or contains(translate(text(),'PIX','pix'),'pix')]/preceding-sibling::span[contains(@class,'price-main')]",
        "//span[contains(@class,'price-main')]",
        "//*[contains(translate(text(),'PIX','pix'),'no pix')]/preceding-sibling::*[contains(text(),'R$')]",
    ]

    for seletor in seletores_pix:
        try:
            elementos = driver.find_elements(By.XPATH, seletor)
            for el in elementos:
                texto = el.text.strip()
                preco = extrair_menor_preco(texto)
                if preco and preco > 5:
                    return preco
        except Exception:
            continue

    # Fallback: maior preço da página que não seja parcela (parcela < preço à vista)
    try:
        todos_elementos = driver.find_elements(By.XPATH, "//*[contains(text(),'R$')]")
        maior_preco = 0.0
        for el in todos_elementos:
            texto = el.text.strip()
            if 'x de' in texto.lower() or 'parcela' in texto.lower():
                continue
            preco = extrair_menor_preco(texto)
            if preco and preco > maior_preco:
                maior_preco = preco
        return maior_preco if maior_preco > 0 else None
    except Exception:
        return None


def testar_outras_variacoes(driver, preco_esperado):
    """
    Quando o preço da variação padrão (cor/tamanho já selecionados ao abrir
    a página) não bate com o esperado, percorre as demais variações do
    produto clicando em cada botão e reconferindo o preço "no Pix", até
    achar uma que bata com o preço esperado.

    Estrutura real (DevTools): cada grupo de variação fica em
    <div id="id.optionN"> ... <div class="buttons-select-container">
    com um <div id="_COR_XXX"> ou <div id="_TAMANHO_XX"> por opção; a opção
    selecionada no momento tem a classe "active" (já foi conferida antes de
    chamar esta função, então é pulada aqui).

    Retorna (preco_encontrado, nome_variacao) se alguma variação tiver o
    preço correto, ou (None, None) se nenhuma variação bater.
    """
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)
    xpath_grupos = "//div[starts-with(@id,'id.option')]"

    try:
        num_grupos = len(driver.find_elements(By.XPATH, xpath_grupos))
    except Exception:
        return None, None

    for indice_grupo in range(num_grupos):
        # XPath posicional (1-based) em vez de guardar a referência do elemento:
        # clicar em uma variação re-renderiza o DOM e invalida referências
        # antigas (StaleElementReferenceException), então tudo é re-localizado
        # do zero a cada tentativa.
        xpath_botoes = (
            f"({xpath_grupos})[{indice_grupo + 1}]"
            "//div[contains(@class,'buttons-select-container')]/div[starts-with(@id,'_')]"
        )

        try:
            num_botoes = len(driver.find_elements(By.XPATH, xpath_botoes))
        except Exception:
            continue

        for indice_botao in range(num_botoes):
            try:
                botoes = driver.find_elements(By.XPATH, xpath_botoes)
                if indice_botao >= len(botoes):
                    continue
                botao = botoes[indice_botao]

                classe = botao.get_attribute('class') or ''
                if 'active' in classe:
                    continue  # é a variação padrão, o preço dela já foi conferido antes

                if (botao.get_attribute('aria-disabled') == 'true') or ('disabled' in classe.lower()):
                    continue  # variação sem estoque/indisponível

                nome_variacao = (botao.get_attribute('id') or '').lstrip('_')

                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", botao)
                driver.execute_script("arguments[0].click();", botao)

                time.sleep(0.8)  # aguarda o preço ser atualizado (troca via AJAX/SPA)
                wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'R$')]")))

                preco_variacao = extrair_preco_pix_na_pagina(driver)
                if preco_variacao is None:
                    continue

                logger.info(f"    Testando variação '{nome_variacao}': R$ {preco_variacao:.2f}")

                if abs(preco_variacao - float(preco_esperado)) <= TOLERANCIA_PRECO:
                    return preco_variacao, nome_variacao

            except Exception as e:
                logger.debug(f"  Falha ao testar variação: {e}")
                continue

    return None, None


# --- PROCESSAMENTO DE CADA PRODUTO ---

def validar_produto(driver, codigo_original, preco_esperado, tentativa=1):
    """
    Pesquisa o produto, extrai o preço no Pix e compara com o preço esperado.

    Retorna dict com: codigo_pesquisado, link, preco_site, status.
    """
    codigo_pesquisado = codigo_original[:-1].strip() if codigo_termina_com_a(codigo_original) else codigo_original

    try:
        encontrou = buscar_produto(driver, codigo_pesquisado)
        if not encontrou:
            return {
                'codigo_pesquisado': codigo_pesquisado,
                'link': None,
                'preco_site': None,
                'status': 'PRODUTO NÃO ENCONTRADO',
                'variacao_correta': None,
            }

        link_produto = selecionar_link_produto(driver)
        if not link_produto:
            return {
                'codigo_pesquisado': codigo_pesquisado,
                'link': None,
                'preco_site': None,
                'status': 'PRODUTO NÃO ENCONTRADO',
                'variacao_correta': None,
            }

        driver.get(link_produto)

        wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'R$')]")))
        except Exception:
            return {
                'codigo_pesquisado': codigo_pesquisado,
                'link': link_produto,
                'preco_site': None,
                'status': 'PÁGINA DO PRODUTO NÃO CARREGOU',
                'variacao_correta': None,
            }

        preco_site = extrair_preco_pix_na_pagina(driver)
        if preco_site is None:
            return {
                'codigo_pesquisado': codigo_pesquisado,
                'link': link_produto,
                'preco_site': None,
                'status': 'PREÇO NÃO ENCONTRADO',
                'variacao_correta': None,
            }

        variacao_correta = None

        if preco_esperado is None or (isinstance(preco_esperado, float) and pd.isna(preco_esperado)):
            status = 'SEM PREÇO ESPERADO NA PLANILHA'
        else:
            diferenca = abs(preco_site - float(preco_esperado))
            status = 'OK' if diferenca <= TOLERANCIA_PRECO else 'DIVERGENTE'

            if status == 'DIVERGENTE':
                logger.info("  Preço padrão divergente. Testando outras cores/tamanhos...")
                preco_variacao, nome_variacao = testar_outras_variacoes(driver, preco_esperado)
                if preco_variacao is not None:
                    preco_site = preco_variacao
                    variacao_correta = nome_variacao
                    status = 'OK (OUTRA VARIAÇÃO)'
                    logger.info(f"  Preço correto encontrado na variação '{nome_variacao}': R$ {preco_site:.2f}")
                else:
                    logger.info("  Nenhuma outra variação com o preço correto. Mantendo DIVERGENTE.")

        return {
            'codigo_pesquisado': codigo_pesquisado,
            'link': link_produto,
            'preco_site': preco_site,
            'status': status,
            'variacao_correta': variacao_correta,
        }

    except Exception as e:
        logger.error(f"  Erro ao validar {codigo_original}: {e}")
        if tentativa < MAX_TENTATIVAS:
            time.sleep(2)
            return validar_produto(driver, codigo_original, preco_esperado, tentativa + 1)
        return {
            'codigo_pesquisado': codigo_pesquisado,
            'link': None,
            'preco_site': None,
            'status': 'ERRO',
            'variacao_correta': None,
        }


# --- ROBÔ PRINCIPAL ---

def iniciar_validacao():
    logger.info("=" * 60)
    logger.info("INICIANDO VALIDAÇÃO DE PREÇOS - SPORTBAY")
    logger.info("=" * 60)

    try:
        logger.info(f"Lendo arquivo: {ARQUIVO_ENTRADA}")
        df = pd.read_excel(ARQUIVO_ENTRADA)
    except FileNotFoundError:
        logger.error(f"Arquivo não encontrado: {ARQUIVO_ENTRADA}")
        return
    except Exception as e:
        logger.error(f"Erro ao ler arquivo: {e}")
        return

    df = df[df.iloc[:, COL_CODIGO].notna()].reset_index(drop=True)
    total_produtos = len(df)
    logger.info(f"Total de produtos a validar: {total_produtos}")

    if total_produtos == 0:
        logger.warning("Nenhum código encontrado na coluna A. Encerrando.")
        return

    opcoes = configurar_chrome()
    driver = None
    try:
        logger.info("Iniciando Google Chrome...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes)
        driver.set_page_load_timeout(20)
        logger.info("Chrome iniciado com sucesso!")
    except Exception as e:
        logger.critical(f"Erro ao tentar abrir o Chrome: {e}")
        return

    resultados = []

    try:
        for index, row in df.iterrows():
            codigo_original = str(row.iloc[COL_CODIGO]).strip()
            produto_nome = row.iloc[COL_PRODUTO] if COL_PRODUTO < len(row) else None
            preco_esperado = row.iloc[COL_PRECO_ESPERADO] if COL_PRECO_ESPERADO < len(row) else None

            if not codigo_original or codigo_original.lower() == 'nan':
                continue

            logger.info(f"\n[{index + 1}/{total_produtos}] Validando código: {codigo_original}")

            resultado = validar_produto(driver, codigo_original, preco_esperado)

            preco_site_fmt = f"R$ {resultado['preco_site']:.2f}" if resultado['preco_site'] is not None else "-"
            preco_esp_fmt = f"R$ {preco_esperado:.2f}" if preco_esperado is not None and not pd.isna(preco_esperado) else "-"
            logger.info(
                f"  Esperado: {preco_esp_fmt} | Site (Pix): {preco_site_fmt} | Status: {resultado['status']}"
            )

            diferenca = None
            if resultado['preco_site'] is not None and preco_esperado is not None and not pd.isna(preco_esperado):
                diferenca = round(resultado['preco_site'] - float(preco_esperado), 2)

            resultados.append({
                'Código': codigo_original,
                'Código Pesquisado': resultado['codigo_pesquisado'],
                'Produto': produto_nome,
                'Preço Esperado': round(float(preco_esperado), 2) if preco_esperado is not None and not pd.isna(preco_esperado) else None,
                'Preço no Site (Pix)': resultado['preco_site'],
                'Diferença (R$)': diferenca,
                'Status': resultado['status'],
                'Variação Correta': resultado.get('variacao_correta'),
                'Link do Produto': resultado['link'],
            })

            time.sleep(TEMPO_ESPERA_ENTRE_PRODUTOS)

    except KeyboardInterrupt:
        logger.warning("\n⚠ Interrupção do usuário detectada!")
    except Exception as e:
        logger.error(f"Erro geral durante o processamento: {e}")
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome fechado")

    if not resultados:
        logger.warning("Nenhum resultado para salvar.")
        return

    df_resultado = pd.DataFrame(resultados)

    try:
        df_resultado.to_excel(ARQUIVO_SAIDA, index=False)
        logger.info(f"\n{'=' * 60}")
        logger.info(f"✓ SUCESSO! Arquivo salvo: {ARQUIVO_SAIDA}")

        contagem = df_resultado['Status'].value_counts()
        for status, qtd in contagem.items():
            logger.info(f"  {status}: {qtd}")
        logger.info(f"{'=' * 60}")
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo de resultado: {e}")


if __name__ == "__main__":
    try:
        iniciar_validacao()
    except Exception as e:
        logger.critical(f"Erro não tratado: {e}")
