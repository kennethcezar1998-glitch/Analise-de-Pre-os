import os
import pandas as pd
import time
import re
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import selenium.webdriver.common.by
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urljoin

# --- CAMINHOS ---
# Pasta deste script = raiz do projeto "Analise de preços site". Planilhas e log
# são resolvidos a partir dela, e não do diretório em que o terminal está.
PASTA_SCRIPT = os.path.dirname(os.path.abspath(__file__))

# --- CONFIGURAÇÕES DE LOGGING ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(PASTA_SCRIPT, 'robo_sportbay.log')),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- CONFIGURAÇÕES ---
NOME_FICHEIRO_ENTRADA = os.path.join(PASTA_SCRIPT, 'Base dos Produtos.xlsx')
NOME_FICHEIRO_SAIDA = os.path.join(PASTA_SCRIPT, 'Resultado_Menor_Preco_Sportbay.xlsx')
TIMEOUT_PAGINA = 15  # segundos para carregar página
TIMEOUT_ELEMENTO = 10  # segundos para encontrar elemento
TEMPO_ESPERA_PADRAO = 3  # espera entre requisições
MAX_TENTATIVAS = 3  # número de tentativas por produto

# --- FUNÇÕES INTELIGENTES ---
def extrair_menor_preco(texto):
    """Extrai o menor preço em formato R$ do texto.
    
    Suporta os formatos brasileiros:
      - R$ 1.149,05  →  1149.05
      - R$ 114,05    →  114.05
      - R$ 1149.05   →  1149.05  (formato inglês, fallback)
    """
    try:
        matches = re.findall(r'R\$\s*([\d\.,]+)', texto)
        menor_valor = float('inf')

        for match in matches:
            try:
                # Formato BR padrão: separador de milhar = '.' e decimal = ','
                # Ex: "1.149,05" ou "114,05"
                if ',' in match:
                    # Remove pontos de milhar, troca vírgula por ponto decimal
                    valor_str = match.replace('.', '').replace(',', '.')
                else:
                    # Sem vírgula: pode ser "1149.05" (inglês) ou "1149" (inteiro)
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

def extrair_codigo_da_url(url):
    """Extrai o código do produto da URL da Sportbay.
    
    Formato da URL: .../CATEGORIA/P-AUSY6MN50M/p
    O '/p' no final é sufixo de rota, não faz parte do código.
    Retorna apenas a penúltima parte, ex: 'P-AUSY6MN50M'
    """
    try:
        if not url:
            return "URL vazia"

        url_limpa = url.split('?')[0].rstrip('/')
        partes = url_limpa.split('/')

        # Ignorar a última parte se for literalmente 'p' (sufixo de rota Sportbay)
        if partes and partes[-1].lower() == 'p':
            partes = partes[:-1]

        if partes:
            return partes[-1]  # Apenas o código, sem /p
        return "Código não identificado"
    except Exception as e:
        logger.warning(f"Erro ao extrair código da URL '{url}': {e}")
        return "Erro na URL"

def codigo_termina_com_a(codigo):
    """Verifica se o código do produto termina com o sufixo 'A' (case-insensitive).

    Alguns códigos da base têm uma variante sem o sufixo 'A' que também é
    válida no site (ex: 'XPTO123A' e 'XPTO123'), então pesquisamos as duas
    formas e ficamos com o menor preço encontrado.
    """
    codigo = (codigo or "").strip()
    return len(codigo) > 1 and codigo[-1].upper() == 'A'

def configurar_chrome():
    """Configura opções do Chrome com máxima compatibilidade"""
    opcoes = Options()
    
    # Opções para estabilidade
    opcoes.add_argument("--start-maximized")
    opcoes.add_argument("--disable-gpu")
    opcoes.add_argument("--no-sandbox")
    opcoes.add_argument("--disable-dev-shm-usage")  # Evita problemas de memória
    opcoes.add_argument("--disable-blink-features=AutomationControlled")  # Anti-bot
    opcoes.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    
    # Desabilitar logs desnecessários
    opcoes.add_argument("--log-level=3")
    
    return opcoes

def extrair_preco_pix_na_pagina(driver):
    """
    Extrai o preço PIX na página do produto.

    Estrutura real (inspecionada via DevTools) dentro do bloco de compra:
        <div class="sc-917c3895-15 kZEWKi">
          <p>
            <span class="price-main">R$&nbsp;80,90</span>
            <span class="priceObs">no <strong>Pix</strong></span>
          </p>
        </div>

    Ou seja: o preço do Pix fica na classe fixa "price-main", dentro do
    mesmo <p> que contém a classe "priceObs" com o texto "Pix". Usamos essa
    relação para não confundir com preços de parcelas.
    """
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    # --- Seletores em ordem de prioridade (do mais específico ao fallback) ---
    seletores_pix = [
        # Estrutura real: <strong>Pix</strong> dentro do mesmo <p> do "price-main"
        "//strong[contains(translate(text(),'PIX','pix'),'pix')]/ancestor::p[1]//span[contains(@class,'price-main')]",
        # Variante caso não seja <p> mas outro container próximo
        "//*[contains(@class,'priceObs')][.//strong[contains(translate(text(),'PIX','pix'),'pix')] or contains(translate(text(),'PIX','pix'),'pix')]/preceding-sibling::span[contains(@class,'price-main')]",
        # Fallback: qualquer "price-main" na página (classe fixa e específica da Sportbay)
        "//span[contains(@class,'price-main')]",
        # Texto "no Pix" próximo ao preço — estrutura comum em sites VTEX
        "//*[contains(translate(text(),'PIX','pix'),'no pix')]/preceding-sibling::*[contains(text(),'R$')]",
        # Span/div com classe de preço de destaque que contenha "Pix"
        "//*[contains(@class,'pix') or contains(@class,'Pix')]//*[contains(text(),'R$')]",
        "//*[contains(@class,'pix') or contains(@class,'Pix') or contains(@class,'PIX')]",
        # Preço dentro do bloco de compra (área lateral direita da página de produto)
        "//div[contains(@class,'buy') or contains(@class,'Buy') or contains(@class,'purchase')]//span[contains(text(),'R$')]",
        # Qualquer elemento que contenha "Pix" e "R$" próximos
        "//*[contains(text(),'Pix') or contains(text(),'PIX') or contains(text(),'pix')]",
    ]

    for seletor in seletores_pix:
        try:
            elementos = driver.find_elements(selenium.webdriver.common.by.By.XPATH, seletor)
            for el in elementos:
                texto = el.text.strip()
                preco = extrair_menor_preco(texto)
                if preco and preco > 5:  # Preço real sempre > R$5 (evita pegar parcelas pequenas)
                    logger.debug(f"  Preço PIX encontrado via seletor específico: R$ {preco:.2f}")
                    return preco
        except Exception:
            continue

    # --- Fallback: pegar o MAIOR preço da página (não a parcela) ---
    # Lógica: parcelas são sempre menores que o preço à vista/PIX
    # então o maior valor presente é o preço real do produto
    logger.debug("  Seletores PIX não encontraram resultado, usando fallback (maior preço)")
    try:
        todos_elementos = driver.find_elements(selenium.webdriver.common.by.By.XPATH, "//*[contains(text(),'R$')]")
        maior_preco = 0.0
        for el in todos_elementos:
            try:
                texto = el.text.strip()
                # Ignorar elementos de parcela que explicitamente mencionam "x de"
                if 'x de' in texto.lower() or 'parcela' in texto.lower():
                    continue
                preco = extrair_menor_preco(texto)
                if preco and preco > maior_preco:
                    maior_preco = preco
            except Exception:
                continue
        return maior_preco if maior_preco > 0 else None
    except Exception as e:
        logger.warning(f"  Fallback de preço falhou: {e}")
        return None


def clicar_menor_preco(driver):
    """
    Marca a opção 'Menor Preço' no filtro de Ordenação da página de resultados.

    Estrutura real (inspecionada via DevTools):
        <h3>Ordenação</h3>
        <div class="sc-d55d46e2-1 cDAieD">
          <label class="sc-d55d46e2-5 ehHWoX">
            <input type="checkbox" class="sc-d55d46e2-8 bdrDld">
            <p class="sc-d55d46e2-10 jHVVBW">Menor Preço</p>
          </label>
          ...
        </div>

    Não é um link, e sim um checkbox dentro de um <label> — clicamos no
    <label> (ou no checkbox, como fallback) em vez de procurar por <a>/<span>.
    As classes com hash (sc-d55d46e2-...) são geradas por build e podem
    mudar, então localizamos pelo texto "Menor Preço" e subimos até o
    <label> ancestral, que é o elemento realmente clicável.

    Retorna True se conseguiu clicar, False caso contrário.
    """
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    # Seletores para a opção "Menor Preço" — em ordem de especificidade
    seletores_menor_preco = [
        # Estrutura real: <p>Menor Preço</p> dentro do <label> clicável
        "//p[normalize-space(text())='Menor Preço']/ancestor::label[1]",
        "//label[.//p[normalize-space(text())='Menor Preço']]",
        # Fallback: clicar direto no checkbox associado
        "//p[normalize-space(text())='Menor Preço']/preceding-sibling::input[@type='checkbox']",
        # Fallbacks genéricos (caso a estrutura mude para link/span)
        "//a[normalize-space(text())='Menor Preço']",
        "//span[normalize-space(text())='Menor Preço']",
        "//li[normalize-space(text())='Menor Preço']",
        "//*[contains(normalize-space(text()),'Menor Preço') and not(contains(@class,'active')) and not(contains(@class,'selected'))]",
        "//*[contains(normalize-space(text()),'Menor Preço')]",
    ]

    for seletor in seletores_menor_preco:
        try:
            elementos = driver.find_elements(selenium.webdriver.common.by.By.XPATH, seletor)
            for el in elementos:
                # Inputs (checkbox) não têm texto visível — o seletor XPath já
                # garante que é o checkbox certo, então pulamos a checagem de texto.
                is_input = el.tag_name.lower() == 'input'
                texto = el.text.strip()
                # Verificar que é exatamente "Menor Preço" e não "Maior Preço"
                if is_input or ('Menor Preço' in texto and 'Maior' not in texto):
                    # Scroll até o elemento para garantir visibilidade
                    driver.execute_script("arguments[0].scrollIntoView(true);", el)
                    time.sleep(0.5)
                    try:
                        el.click()
                    except Exception:
                        # Elemento pode estar coberto/oculto por CSS custom — clique via JS
                        driver.execute_script("arguments[0].click();", el)
                    logger.debug("  → Marcou 'Menor Preço'")

                    # Aguardar página recarregar com nova ordenação
                    wait.until(
                        EC.presence_of_all_elements_located((selenium.webdriver.common.by.By.XPATH, "//a[contains(@href, '/p')]"))
                    )
                    time.sleep(1)  # Pequena pausa extra para estabilizar DOM
                    return True
        except Exception as e:
            logger.debug(f"  Seletor '{seletor}' falhou: {e}")
            continue

    logger.warning("  ⚠ Não encontrou opção 'Menor Preço' — usando ordenação padrão")
    return False


def buscar_produto(driver, codigo_produto):
    """
    Realiza a busca pelo código do produto usando a barra de busca real do site.

    Estrutura real (inspecionada via DevTools):
        <button type="button" aria-label="search_button" id="seach_button">
        <input id="search_user_component" name="searchComponent"
               placeholder="O que você procura?" type="text">

    Digitar no campo e clicar no botão é mais confiável do que só montar a
    URL com querystring, pois garante que a busca é disparada pelo próprio
    JS da Sportbay (SPA), replicando o comportamento real do usuário.
    Se a interação com a UI falhar por algum motivo, cai para a URL direta.

    Retorna True se encontrou resultados, False caso contrário.
    """
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    try:
        driver.get("https://www.sportbay.com.br/")

        campo_busca = wait.until(
            EC.element_to_be_clickable((selenium.webdriver.common.by.By.ID, "search_user_component"))
        )
        campo_busca.clear()
        campo_busca.send_keys(codigo_produto)

        # Clicar no botão da lupa (id="seach_button") para disparar a busca.
        # O botão contém um <svg> interno, então usamos wait "clickable" e,
        # se ainda assim o clique nativo for interceptado/ignorado, forçamos
        # via JS diretamente no elemento do botão.
        try:
            botao_busca = wait.until(
                EC.element_to_be_clickable((selenium.webdriver.common.by.By.ID, "seach_button"))
            )
            botao_busca.click()
        except Exception as e_click:
            logger.debug(f"  Clique nativo no botão de busca falhou ({e_click}), tentando via JS...")
            botao_busca = driver.find_element(selenium.webdriver.common.by.By.ID, "seach_button")
            driver.execute_script("arguments[0].click();", botao_busca)

        # IMPORTANTE: a home já tem produtos com links '/p' (carrossel "Oferta
        # do Dia"), então checar só a presença de links '/p' passa mesmo sem a
        # busca ter sido disparada — e o robô acaba pegando o primeiro produto
        # da home. Por isso esperamos a URL realmente mudar para "/busca"
        # antes de considerar a busca concluída.
        try:
            wait.until(EC.url_contains("busca"))
        except Exception:
            # Clique no botão não navegou para a busca — tenta Enter no campo como último recurso
            logger.debug("  → Clique no botão não navegou para a busca, tentando Enter no campo")
            campo_busca.send_keys(Keys.RETURN)
            wait.until(EC.url_contains("busca"))

        wait.until(
            EC.presence_of_all_elements_located((selenium.webdriver.common.by.By.XPATH, "//a[contains(@href, '/p')]"))
        )

        logger.debug("  → Busca realizada via barra de pesquisa")
        return True
    except Exception as e:
        logger.debug(f"  Busca via barra de pesquisa falhou ({e}), tentando via URL direta...")
        try:
            url_busca = f"https://www.sportbay.com.br/busca?q={codigo_produto}&O=OrderByPricASC"
            driver.get(url_busca)
            wait.until(
                EC.presence_of_all_elements_located((selenium.webdriver.common.by.By.XPATH, "//a[contains(@href, '/p')]"))
            )
            return True
        except Exception:
            return False


def _buscar_e_extrair_preco(driver, codigo_busca):
    """
    Executa uma busca completa por um único código: pesquisa na Sportbay,
    ordena por menor preço, entra no primeiro resultado e extrai o preço PIX.

    Retorna tupla (código_site, preço_pix_float). Qualquer um dos dois pode
    vir None se a etapa correspondente falhar (produto não encontrado,
    página não carregou, preço não localizado, etc).
    """
    wait = WebDriverWait(driver, TIMEOUT_ELEMENTO)

    # 1. BUSCAR PRODUTO — via barra de busca real do site (com fallback para URL direta)
    encontrou = buscar_produto(driver, codigo_busca)
    if not encontrou:
        logger.warning(f"  → Nenhum produto encontrado para: {codigo_busca}")
        return (None, None)

    # 2. CLICAR EM "MENOR PREÇO" PARA GARANTIR ORDENAÇÃO CORRETA
    clicou = clicar_menor_preco(driver)
    if clicou:
        logger.debug("  → Resultados reordenados por menor preço")
    else:
        logger.warning("  → Continuando sem ordenar por menor preço")

    # 3. PEGAR URL DO PRIMEIRO RESULTADO (agora é o mais barato)
    links_resultado = driver.find_elements(selenium.webdriver.common.by.By.XPATH, "//a[contains(@href, '/p')]")

    url_produto = None
    codigo_site = None
    for link in links_resultado:
        href = link.get_attribute('href') or ''
        # URL de produto da Sportbay termina com /p
        if href.endswith('/p') or '/p?' in href:
            url_produto = href.split('?')[0]  # Remove query string
            codigo_site = extrair_codigo_da_url(url_produto)
            break

    if not url_produto:
        logger.warning(f"  → Nenhum link de produto válido encontrado")
        return (None, None)

    logger.debug(f"  Navegando para produto: {url_produto}")

    # 4. NAVEGAR ATÉ A PÁGINA DO PRODUTO
    driver.get(url_produto)

    try:
        # Aguardar o bloco de preço carregar
        wait.until(
            EC.presence_of_element_located((selenium.webdriver.common.by.By.XPATH, "//*[contains(text(),'R$')]"))
        )
    except:
        logger.warning(f"  → Página do produto não carregou corretamente")
        return (codigo_site, None)

    # 5. EXTRAIR PREÇO PIX
    preco_pix = extrair_preco_pix_na_pagina(driver)

    if preco_pix and preco_pix > 0:
        logger.info(f"  ✓ Preço PIX: R$ {preco_pix:.2f} | Código: {codigo_site}")
        return (codigo_site, preco_pix)
    else:
        logger.warning(f"  → Não foi possível extrair preço PIX")
        return (codigo_site, None)


def processar_produto(driver, codigo_produto, tentativa=1):
    """
    Processa um produto individual com retry logic.

    Se o código termina com 'A', pesquisa também a variante sem o sufixo
    (ex: 'XPTO123A' → também busca 'XPTO123') e compara os preços das duas
    buscas, retornando apenas o resultado com o MENOR preço.

    Retorna tupla (código_site, preço_pix)
    """
    try:
        logger.info(f"[Tentativa {tentativa}/{MAX_TENTATIVAS}] Pesquisando: {codigo_produto}")

        # 1. Busca com o código original
        resultados = [_buscar_e_extrair_preco(driver, codigo_produto)]

        # 2. Se o código termina com 'A', pesquisa também sem o sufixo
        if codigo_termina_com_a(codigo_produto):
            codigo_sem_a = codigo_produto[:-1].strip()
            logger.info(f"  → Código termina com 'A' — pesquisando também sem o sufixo: {codigo_sem_a}")
            time.sleep(TEMPO_ESPERA_PADRAO)
            resultados.append(_buscar_e_extrair_preco(driver, codigo_sem_a))

        # 3. Escolher o resultado com o MENOR preço válido entre as buscas
        validos = [(cs, p) for cs, p in resultados if p is not None]
        if validos:
            codigo_site_final, preco_final = min(validos, key=lambda item: item[1])
            logger.info(f"  ✓ Sucesso! Menor preço entre as buscas: R$ {preco_final:.2f} | Código: {codigo_site_final}")
            return (codigo_site_final, f"{preco_final:.2f}")

        # Nenhuma busca retornou um preço válido
        codigo_site_fallback = next((cs for cs, _ in resultados if cs), None)
        if codigo_site_fallback:
            logger.warning(f"  → Produto encontrado, mas não foi possível extrair o preço PIX")
            return (codigo_site_fallback, "Erro")

        logger.warning(f"  → Nenhum produto encontrado para: {codigo_produto}")
        return ("N/A", "N/A")

    except Exception as e:
        logger.error(f"  ✗ Erro ao processar {codigo_produto}: {str(e)}")

        if tentativa < MAX_TENTATIVAS:
            time.sleep(2)
            return processar_produto(driver, codigo_produto, tentativa + 1)
        else:
            return ("Erro", "Erro")

# --- O ROBÔ PRINCIPAL ---
def iniciar_robo():
    """Função principal de automação"""
    logger.info("=" * 60)
    logger.info("INICIANDO ROBÔ DE SCRAPING - SPORTBAY")
    logger.info("=" * 60)
    
    # 1. LER DADOS DE ENTRADA
    try:
        logger.info(f"Lendo arquivo: {NOME_FICHEIRO_ENTRADA}")
        df = pd.read_excel(NOME_FICHEIRO_ENTRADA)
        logger.info(f"Total de produtos a processar: {len(df)}")
    except FileNotFoundError:
        logger.error(f"Arquivo não encontrado: {NOME_FICHEIRO_ENTRADA}")
        return
    except Exception as e:
        logger.error(f"Erro ao ler arquivo: {e}")
        return
    
    # Validar coluna
    coluna_busca = df.columns[0]
    logger.info(f"Usando coluna: '{coluna_busca}'")
    
    # 2. INICIALIZAR SELENIUM
    codigos_site_encontrados = []
    menores_precos = []
    
    opcoes = configurar_chrome()
    
    driver = None
    try:
        logger.info("Iniciando Google Chrome...")
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=opcoes
        )
        driver.set_page_load_timeout(TIMEOUT_PAGINA)
        logger.info("Chrome iniciado com sucesso!")
        
    except Exception as e:
        logger.critical(f"Erro ao tentar abrir o Chrome: {e}")
        return
    
    # 3. PROCESSAR CADA PRODUTO
    total_produtos = len(df)
    
    try:
        for index, row in df.iterrows():
            codigo_produto = str(row[coluna_busca]).strip()
            
            # Validar código
            if not codigo_produto or codigo_produto.lower() == 'nan':
                logger.warning(f"[{index + 1}/{total_produtos}] Código vazio ou inválido")
                codigos_site_encontrados.append("N/A")
                menores_precos.append("N/A")
                continue
            
            logger.info(f"\n[{index + 1}/{total_produtos}] Processando...")
            
            # Processar produto
            codigo_site, menor_preco = processar_produto(driver, codigo_produto)
            codigos_site_encontrados.append(codigo_site)
            menores_precos.append(menor_preco)
            
            # Esperar entre requisições para não sobrecarregar servidor
            time.sleep(TEMPO_ESPERA_PADRAO)
    
    except KeyboardInterrupt:
        logger.warning("\n⚠ Interrupção do usuário detectada!")
    
    except Exception as e:
        logger.error(f"Erro geral durante processamento: {e}")
    
    finally:
        # Fechar navegador
        if driver:
            driver.quit()
            logger.info("Chrome fechado")
    
    # 4. SALVAR RESULTADOS
    try:
        df['Código Site Sportbay'] = codigos_site_encontrados
        df['Menor Preço'] = menores_precos
        
        df.to_excel(NOME_FICHEIRO_SAIDA, index=False)
        logger.info(f"\n{'=' * 60}")
        logger.info(f"✓ SUCESSO! Arquivo salvo: {NOME_FICHEIRO_SAIDA}")
        logger.info(f"Total processado: {len(df)}")
        logger.info(f"{'=' * 60}")
        
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo de resultado: {e}")

if __name__ == "__main__":
    try:
        iniciar_robo()
    except Exception as e:
        logger.critical(f"Erro não tratado: {e}")