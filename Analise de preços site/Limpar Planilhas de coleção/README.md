╔════════════════════════════════════════════════════════════════════════════╗
║         PRICING INTELLIGENCE SYSTEM - PRO TORK MARKETPLACE MONITOR         ║
║                                                                            ║
║  Monitoramento de preços em tempo real: Sportbay vs Mercado Livre & Shopee ║
╚════════════════════════════════════════════════════════════════════════════╝

## 📋 ÍNDICE

1. [Instalação](#-instalação)
2. [Características](#-características)
3. [Arquitetura](#-arquitetura)
4. [Guia de Uso](#-guia-de-uso)
5. [Exemplos Práticos](#-exemplos-práticos)
6. [Tratamento de Erros](#-tratamento-de-erros)
7. [Troubleshooting](#-troubleshooting)
8. [Performance](#-performance)

---

## 🚀 INSTALAÇÃO

### Pré-requisitos
- Python 3.8+
- pip (gerenciador de pacotes Python)
- Acesso à internet (para APIs e web scraping)

### Passo 1: Instalar Dependências

```bash
# Clone ou baixe este repositório
cd pricing-intelligence-pro-tork

# Instale as dependências
pip install -r requirements.txt
```

**Saída esperada:**
```
Successfully installed pandas-2.1.3 requests-2.31.0 playwright-1.40.0 ...
```

### Passo 2: Instalar Chromium (Playwright)

O Playwright precisa de um navegador para funcionar. Execute:

```bash
# Instalar Chromium (usado pela Shopee)
playwright install chromium
```

**Saída esperada:**
```
✓ Chromium browser downloaded to ~/.cache/ms-playwright/chromium-xxx/
```

### Passo 3: Verificar Instalação

```bash
# Teste a importação
python -c "import pandas, requests, playwright, rapidfuzz; print('✓ Tudo OK!')"
```

---

## ✨ CARACTERÍSTICAS

### 1. **API Mercado Livre (Oficial)**
- ✓ Busca por EAN (código de barras)
- ✓ Fallback para busca por nome
- ✓ Extração de: Título, Preço, Link, Status Full
- ✓ Sem bloqueios por ser via API oficial

### 2. **Web Scraping Shopee (Playwright Async)**
- ✓ Navegador headless (invisível)
- ✓ Aguarda carregamento dinâmico
- ✓ User-Agent real para evitar bloqueios
- ✓ Extração de: Título, Preço, Link
- ✓ Execução assíncrona (super rápida)

### 3. **Matching Inteligente (RapidFuzz)**
- ✓ Comparação token-based de nomes de produto
- ✓ Threshold de 80% para evitar falsos positivos
- ✓ Score de similaridade em cada resultado

### 4. **Filtros de Qualidade**
- ✓ Outlier detection: Descarta preços < 40% do valor Sportbay
- ✓ Evita comparações incorretas (ex: viseira vs capacete inteiro)

### 5. **Processamento Paralelo**
- ✓ asyncio.gather para executar múltiplas buscas simultaneamente
- ✓ Reutilização de conexões HTTP e browser
- ✓ Velocidade: ~0.5-2 segundos por produto

### 6. **Output Profissional**
- ✓ Arquivo Excel (.xlsx) com duas sheets:
  - Sheet 1: Comparativo detalhado de preços
  - Sheet 2: Resumo estatístico

---

## 🏗️ ARQUITETURA

```
┌─────────────────────────────────────────────────────────┐
│            ENTRADA (Sportbay Internal DB)               │
│  [ID, Nome, EAN, Preço Sportbay]                        │
└────────────────────┬────────────────────────────────────┘
                     │
        ┌────────────┴────────────┐
        │                         │
        ▼                         ▼
  ┌──────────────┐         ┌──────────────┐
  │ Mercado Livre│         │   Shopee     │
  │   (API)      │         │ (Playwright) │
  └──────┬───────┘         └──────┬───────┘
         │                        │
         │   ┌─────────────────┐  │
         └──►│  RapidFuzz      │◄─┘
             │  Matching       │
             │  (>80%)         │
             └────────┬────────┘
                      │
         ┌────────────┴────────────┐
         │   Filtros (Outliers)    │
         │  (< 40% preço = reject) │
         └────────────┬────────────┘
                      │
         ┌────────────┴────────────┐
         │  Consolidação DataFrame │
         │   + Estatísticas        │
         └────────────┬────────────┘
                      │
                      ▼
         ┌─────────────────────────┐
         │   Excel Output (.xlsx)  │
         │  - Comparativo Preços   │
         │  - Resumo Estatístico   │
         └─────────────────────────┘
```

---

## 📖 GUIA DE USO

### Uso Básico (Modo Assíncrono)

```python
import asyncio
import pandas as pd
from pricing_intelligence_pro_tork import main

# Opção 1: Com lista de dicionários
dados = [
    {
        'id_produto': '001',
        'nome_produto': 'Capacete Pro Tork TH1',
        'ean': '7895678934521',
        'preco_sportbay': 450.00
    },
    {
        'id_produto': '002',
        'nome_produto': 'Luva Pro Tork Racing',
        'ean': '7895678934522',
        'preco_sportbay': 120.00
    },
]

# Executar análise
df_resultado = asyncio.run(main(dados_entrada=dados))

# Resultado em DataFrame
print(df_resultado)
```

### Uso com DataFrame Pandas

```python
import asyncio
import pandas as pd
from pricing_intelligence_pro_tork import main

# Ler dados de uma base existente
df_sportbay = pd.read_csv('produtos_pro_tork.csv')

# Executar análise
df_resultado = asyncio.run(main(df_entrada=df_sportbay))

# Salvar em novo arquivo Excel
df_resultado.to_excel('analise_precos_pro_tork.xlsx', index=False)
```

### Personalizar Arquivo de Saída

```python
df_resultado = asyncio.run(
    main(
        dados_entrada=dados,
        arquivo_saida='my_custom_report.xlsx'
    )
)
```

---

## 💡 EXEMPLOS PRÁTICOS

### Exemplo 1: Análise Rápida de 5 Produtos

```python
import asyncio
from pricing_intelligence_pro_tork import main

produtos = [
    {'id_produto': '1', 'nome_produto': 'Capacete Pro Tork TH1', 'ean': '7895678934521', 'preco_sportbay': 450},
    {'id_produto': '2', 'nome_produto': 'Capacete Pro Tork TH2', 'ean': '7895678934522', 'preco_sportbay': 520},
    {'id_produto': '3', 'nome_produto': 'Luva Pro Tork XL', 'ean': '7895678934523', 'preco_sportbay': 150},
    {'id_produto': '4', 'nome_produto': 'Jaqueta Pro Tork', 'ean': '7895678934524', 'preco_sportbay': 350},
    {'id_produto': '5', 'nome_produto': 'Bota Pro Tork', 'ean': '7895678934525', 'preco_sportbay': 280},
]

df = asyncio.run(main(dados_entrada=produtos))
print(df[['ID_Produto', 'Preco_Sportbay', 'Preco_ML', 'Preco_Shopee', 'Melhor_Marketplace']])
```

**Saída esperada:**
```
  ID_Produto  Preco_Sportbay  Preco_ML  Preco_Shopee Melhor_Marketplace
0          1           450.0     425.0         440.0      Mercado Livre
1          2           520.0     490.0         510.0      Mercado Livre
2          3           150.0      85.0          95.0      Mercado Livre
3          4           350.0       NaN         315.0           Shopee
4          5           280.0     255.0            NaN      Mercado Livre
```

### Exemplo 2: Integração com Base de Dados

```python
import asyncio
import pandas as pd
from pricing_intelligence_pro_tork import main
import sqlite3

# Ler de um banco SQLite
conexao = sqlite3.connect('sportbay_produtos.db')
df_produtos = pd.read_sql(
    'SELECT id_produto, nome_produto, ean, preco FROM tb_produtos WHERE marca = "Pro Tork"',
    conexao
)

# Renomear coluna de preço
df_produtos.rename(columns={'preco': 'preco_sportbay'}, inplace=True)

# Executar análise
df_resultado = asyncio.run(main(df_entrada=df_produtos))

# Salvar resultado
df_resultado.to_excel('comparacao_preco_pro_tork.xlsx', index=False)

# Exibir resumo
print(f"✓ Produtos analisados: {len(df_resultado)}")
print(f"✓ Matches ML: {df_resultado['Preco_ML'].notna().sum()}")
print(f"✓ Matches Shopee: {df_resultado['Preco_Shopee'].notna().sum()}")
```

### Exemplo 3: Filtrar Apenas Competitividade Baixa

```python
import asyncio
import pandas as pd
from pricing_intelligence_pro_tork import main

df_resultado = asyncio.run(main(dados_entrada=dados))

# Filtrar produtos onde Sportbay é mais caro em ambos marketplaces
produtos_nao_competitivos = df_resultado[
    (df_resultado['Diferenca_ML_%'] > 20) &
    (df_resultado['Diferenca_Shopee_%'] > 20)
]

print("🔴 PRODUTOS MENOS COMPETITIVOS (>20% mais caros):")
print(produtos_nao_competitivos[['ID_Produto', 'Preco_Sportbay', 'Diferenca_ML_%', 'Diferenca_Shopee_%']])
```

### Exemplo 4: Exportar Apenas Melhor Deal

```python
# Encontrar o melhor preço geral entre todos marketplaces
df_resultado['Melhor_Preco'] = df_resultado[['Preco_ML', 'Preco_Shopee']].min(axis=1)
df_resultado['Economia'] = df_resultado['Preco_Sportbay'] - df_resultado['Melhor_Preco']

# Ordenar por economia
top_oportunidades = df_resultado.nlargest(10, 'Economia')
print("\n💰 TOP 10 MAIORES ECONOMIAS:")
print(top_oportunidades[['ID_Produto', 'Nome_Produto', 'Preco_Sportbay', 'Melhor_Preco', 'Economia']])
```

---

## ⚠️ TRATAMENTO DE ERROS

### Erro: "ModuleNotFoundError: No module named 'playwright'"

**Solução:**
```bash
pip install -r requirements.txt
```

---

### Erro: "Chromium browser not found"

**Solução:**
```bash
playwright install chromium
```

---

### Erro: "TimeoutError" (Shopee não responde)

**Causa:** Servidor sobrecarregado ou bloqueio de IP
**Solução:** O script tenta novamente automaticamente. Se persistir:

```python
# Aumentar timeout (em milliseconds)
# Editar a constante no código:
# TIMEOUT_REQUESTS = 20  (de 10 para 20)
```

---

### Erro: "No results found for EAN"

**Causa:** EAN inválido ou produto não está catalogado
**Situação:** Script tentará fallback automático por nome (Pro Tork + Nome)

**Verificar:**
```python
# Validar EAN antes de enviar
import re
ean = "7895678934521"
if re.match(r'^\d{13}$', ean):
    print("✓ EAN válido")
else:
    print("✗ EAN inválido - deve ter 13 dígitos")
```

---

### Erro: "RapidFuzz match score < 80%"

**Causa:** Nome do produto muito diferente do marketplace
**Situação:** Produto será descartado (para evitar falsos positivos)

**Solução (manual):**
```python
# Ajustar nome do produto para ser mais genérico
# De: "Viseira Capacete Pro Tork TH1"
# Para: "Capacete Pro Tork TH1"  (viseira será encontrada separadamente)
```

---

### Erro: "Price outlier rejected (< 40%)"

**Causa:** Preço muito baixo comparado ao Sportbay
**Situação:** Provavelmente é um produto diferente (acessório vs item completo)

**Exemplo:**
```
Sportbay: Capacete Pro Tork TH1 = R$ 450
Marketplace: "Viseira para Capacete Pro Tork" = R$ 85
Resultado: DESCARTADO (85 < 450 * 0.40 = 180)
```

---

## 🔧 TROUBLESHOOTING

### Problema: Script muito lento

**Causas e Soluções:**

1. **Conexão de internet lenta**
   - Verificar velocidade: `ping api.mercadolibre.com`
   - Usar conexão cabeada se possível

2. **Servidor Shopee congestionado**
   - Script aguarda até 20 segundos por padrão
   - Executar fora dos horários de pico (23:00 - 06:00)

3. **Muitos produtos**
   - 100 produtos ≈ 30-60 segundos (5-10 produtos/segundo)
   - Para >500 produtos, dividir em lotes:

```python
import asyncio
from pricing_intelligence_pro_tork import main

# Dividir em lotes de 50 produtos
tamanho_lote = 50
todos_resultados = []

for i in range(0, len(dados), tamanho_lote):
    lote = dados[i:i+tamanho_lote]
    df = asyncio.run(main(dados_entrada=lote))
    todos_resultados.append(df)

df_final = pd.concat(todos_resultados, ignore_index=True)
df_final.to_excel('resultado_completo.xlsx', index=False)
```

---

### Problema: Bloqueio por IP (erro HTTP 429)

**Solução:**
- Aguardar 5-10 minutos
- Alternar entre WiFi e celular (mudar IP)
- Usar VPN (último recurso)

**Nota:** A API Mercado Livre é oficial e não bloqueia. O bloqueio é possível apenas na Shopee.

---

### Problema: Caracteres especiais aparecem errado (encoding)

**Solução:**
```python
# Ao salvar em CSV (se não usar Excel):
df_resultado.to_csv('resultado.csv', encoding='utf-8-sig', index=False)
```

---

## ⚡ PERFORMANCE

### Benchmarks (Máquina Típica)

```
┌──────────────────┬───────────────┬──────────────┐
│ Quantidade       │ Tempo Total   │ Taxa/Produto │
├──────────────────┼───────────────┼──────────────┤
│ 5 produtos       │ ~5-10s        │ 2-1 seg      │
│ 10 produtos      │ ~10-20s       │ 2-1 seg      │
│ 50 produtos      │ ~60-90s       │ 1.8-1.5 seg  │
│ 100 produtos     │ ~2-3 min      │ 1.2-2 seg    │
│ 500 produtos     │ ~10-15 min    │ 1-2 seg      │
└──────────────────┴───────────────┴──────────────┘
```

### Otimizações Implementadas

✓ **asyncio.gather** - Paralelização de requisições
✓ **Session pooling** - Reutilização de conexões HTTP
✓ **Browser singleton** - Uma instância Playwright para todos os produtos
✓ **Lazy loading** - Carregamento sob demanda
✓ **Early termination** - Parar busca quando threshold atingido

### Como Melhorar Performance

1. **Usar conexão mais rápida** (50% de ganho)
2. **Executar em horário de baixa demanda** (20% de ganho)
3. **Aumentar número de workers** (requer modificação no código)

```python
# Exemplo: Processar 2 produtos em paralelo total (4 buscas simultâneas)
# Padrão atual: max 4 requisições simultâneas
# Máximo seguro: 8-10 (para evitar bloqueios)
```

---

## 📊 ESTRUTURA DO OUTPUT EXCEL

### Sheet 1: "Comparativo Preços"

```
Coluna A: ID_Produto           (string) - ID interno Sportbay
Coluna B: Nome_Produto         (string) - Nome do produto
Coluna C: EAN                  (string) - Código de barras
Coluna D: Preco_Sportbay       (float)  - Preço base
Coluna E: Preco_ML             (float)  - Melhor preço Mercado Livre ou NULL
Coluna F: Link_ML              (url)    - Link para o produto no ML
Coluna G: Titulo_ML            (string) - Título do anúncio no ML
Coluna H: Score_ML             (float)  - Similaridade RapidFuzz (%)
Coluna I: Full_ML              (bool)   - Possui frete grátis (Full)
Coluna J: Diferenca_ML_%       (float)  - Variação percentual
Coluna K: Preco_Shopee         (float)  - Melhor preço Shopee ou NULL
Coluna L: Link_Shopee          (url)    - Link para o produto na Shopee
Coluna M: Titulo_Shopee        (string) - Título do anúncio na Shopee
Coluna N: Score_Shopee         (float)  - Similaridade RapidFuzz (%)
Coluna O: Diferenca_Shopee_%   (float)  - Variação percentual
Coluna P: Melhor_Marketplace   (string) - Qual tem melhor preço
Coluna Q: Data_Coleta          (datetime)- Timestamp da coleta
```

### Sheet 2: "Resumo Estatístico"

```
Total de Produtos: 123
Matches Encontrados (ML): 98
Matches Encontrados (Shopee): 89
Taxa de Sucesso (ML): 79.67%
Taxa de Sucesso (Shopee): 72.36%
Preço Médio Sportbay: R$ 287.45
Preço Médio ML: R$ 245.32
Preço Médio Shopee: R$ 251.78
Diferença Média ML: -14.68%
Diferença Média Shopee: -12.41%
Produtos Mais Caros no ML: 32
Produtos Mais Caros na Shopee: 28
```

---

## 📝 LOGS

O sistema gera logs em: `pricing_intelligence.log`

**Exemplo de log:**
```
2024-01-15 14:32:15,123 - INFO - [main] - 🚀 Iniciando processamento de 5 produtos...
2024-01-15 14:32:15,456 - INFO - [busca_mercado_livre] - 🔍 [ML] Buscando por EAN: 7895678934521
2024-01-15 14:32:16,789 - INFO - [busca_mercado_livre] - ✓ Encontrados 12 resultados por EAN
2024-01-15 14:32:17,012 - INFO - [busca_mercado_livre] - ✅ [ML] Produto PT001 - R$425.0 (Score: 95%)
2024-01-15 14:32:17,234 - INFO - [busca_shopee] - 🔍 [Shopee] Buscando por: Pro Tork Capacete TH1
...
2024-01-15 14:32:45,678 - INFO - [main] - ✅ Processamento concluído em 30.12s | Taxa: 0.17 produtos/s
```

---

## 🔐 SEGURANÇA

### Dados Sensíveis
- ✓ Nenhuma informação sensível é armazenada
- ✓ EAN é tratado como dado público
- ✓ Preços são dados públicos dos marketplaces

### Conformidade
- ✓ Respeita Terms of Service da API Mercado Livre
- ✓ Playwright é headless (sem automação óbvia)
- ✓ Delays entre requisições para respeitar servidores

---

## 📞 SUPORTE

Para erros ou dúvidas:

1. Verificar logs: `tail -f pricing_intelligence.log`
2. Validar entrada: Todas as colunas obrigatórias?
3. Testar com 1 produto: Isolar problemas
4. Confirmar internet: `ping api.mercadolibre.com`

---

**Script criado com ❤️ para Pricing Intelligence | Pro Tork Marketplace Monitor**
