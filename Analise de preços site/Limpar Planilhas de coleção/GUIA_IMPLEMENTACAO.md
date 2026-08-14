╔════════════════════════════════════════════════════════════════════════════╗
║     GUIA PRÁTICO DE IMPLEMENTAÇÃO - PRICING INTELLIGENCE PRO TORK          ║
║                         Sportbay × ML × Shopee                            ║
╚════════════════════════════════════════════════════════════════════════════╝

## 🎯 SITUAÇÃO ATUAL

✓ Você tem uma planilha com 9 produtos Pro Tork
✓ Colunas: id_produto, nome_produto, ean, preco_sportbay
✓ Pronta para rodar a análise de preços

---

## 📋 PASSO 1: PREPARAR O AMBIENTE (5 MINUTOS)

### 1.1 Criar Pasta do Projeto

```bash
# No seu terminal (Windows, Mac ou Linux):
mkdir meu_projeto_pricing
cd meu_projeto_pricing
```

### 1.2 Copiar Arquivos para a Pasta

Copie estes 4 arquivos para a pasta:
1. `pricing_intelligence_pro_tork.py` (Script principal)
2. `requirements.txt` (Dependências)
3. `setup_helper.py` (Helper de setup)
4. `analise_de_preços.xlsx` (Sua planilha)

Estrutura final:
```
meu_projeto_pricing/
├── pricing_intelligence_pro_tork.py
├── requirements.txt
├── setup_helper.py
├── analise_de_preços.xlsx
└── README.md
```

### 1.3 Instalar Dependências (10 MINUTOS)

```bash
# Opção A: Automático (RECOMENDADO)
python setup_helper.py install

# Saída esperada:
# ✓ Dependências instaladas com sucesso!
# ✓ Chromium instalado com sucesso!
```

**OU**

```bash
# Opção B: Manual
pip install -r requirements.txt
playwright install chromium
```

### 1.4 Testar Instalação (5 MINUTOS)

```bash
# Verificar se tudo está OK
python setup_helper.py test

# Saída esperada:
# ✓ Pandas (manipulação de dados)
# ✓ Requests (HTTP)
# ✓ Playwright (scraping)
# ✓ RapidFuzz (string matching)
# ✓ OpenPyXL (Excel)
# ✓ API respondeu com sucesso
# ✓ Playwright inicializado com sucesso
```

---

## 🚀 PASSO 2: RODAR ANÁLISE COM SUA PLANILHA (2 MINUTOS)

### 2.1 Script Rápido (Recomendado)

Crie um arquivo chamado `executar_analise.py` na pasta do projeto:

```python
"""
Script rápido para analisar sua planilha de preços
"""

import asyncio
import pandas as pd
from pricing_intelligence_pro_tork import main

# OPÇÃO A: Carregar de arquivo Excel (MAIS FÁCIL)
print("📂 Carregando planilha 'analise_de_preços.xlsx'...")
df_entrada = pd.read_excel('analise_de_preços.xlsx')

print(f"📦 {len(df_entrada)} produtos carregados:")
print(df_entrada[['id_produto', 'nome_produto', 'preco_sportbay']].to_string())

# Executar análise
print("\n⏳ Iniciando análise (pode demorar 1-2 minutos)...\n")
df_resultado = asyncio.run(
    main(
        df_entrada=df_entrada,
        arquivo_saida='resultado_analise_precos.xlsx'
    )
)

# Exibir resultados
print("\n" + "="*80)
print("📊 RESULTADOS:")
print("="*80 + "\n")

# Resumo rápido
print("📈 RESUMO GERAL:")
print(f"   Total de produtos: {len(df_resultado)}")
print(f"   Matches Mercado Livre: {df_resultado['Preco_ML'].notna().sum()}")
print(f"   Matches Shopee: {df_resultado['Preco_Shopee'].notna().sum()}")

# Tabela com resultados principais
print("\n💰 PREÇOS E DIFERENÇAS:")
tabela = df_resultado[[
    'ID_Produto',
    'Preco_Sportbay',
    'Preco_ML',
    'Diferenca_ML_%',
    'Preco_Shopee',
    'Diferenca_Shopee_%',
    'Melhor_Marketplace'
]].copy()

print(tabela.to_string(index=False))

# Encontrar melhores deals
print("\n💎 MELHORES OPORTUNIDADES (Maior diferença de preço):")
df_resultado['Economia_Potencial'] = df_resultado[[
    'Diferenca_ML_%',
    'Diferenca_Shopee_%'
]].min(axis=1)

top_oportunidades = df_resultado.nlargest(5, 'Economia_Potencial')
for _, row in top_oportunidades.iterrows():
    print(f"\n   {row['Nome_Produto']}")
    print(f"   Preço Sportbay: R${row['Preco_Sportbay']:.2f}")
    print(f"   Melhor em {row['Melhor_Marketplace']}: R${min(row['Preco_ML'], row['Preco_Shopee']):.2f}")
    print(f"   Economia: {abs(row['Economia_Potencial']):.1f}%")

print("\n✅ Análise salva em: resultado_analise_precos.xlsx")
print("="*80 + "\n")
```

### 2.2 Executar o Script

```bash
# Windows
python executar_analise.py

# Mac/Linux
python3 executar_analise.py
```

**Saída esperada (primeiras linhas):**
```
📂 Carregando planilha 'analise_de_preços.xlsx'...
📦 9 produtos carregados:
     id_produto                                            nome_produto  preco_sportbay
0  P-36ICMN7CJ5              CAPACETE ABERTO PRO TORK NEW LIBERTY 3         127.48
1  P-6Y2SPRR4G4            VISEIRA CAPACETE MIXS MX2 GLADIATOR FUMÊ          85.91
...

⏳ Iniciando análise (pode demorar 1-2 minutos)...

[LOGS DETALHADOS]

📊 RESULTADOS:
================================================================================

📈 RESUMO GERAL:
   Total de produtos: 9
   Matches Mercado Livre: 7
   Matches Shopee: 6

💰 PREÇOS E DIFERENÇAS:
ID_Produto   Preco_Sportbay  Preco_ML  Diferenca_ML_%  Preco_Shopee  Diferenca_Shopee_%  Melhor_Marketplace
P-36ICMN7CJ5          127.48     95.00          -25.5         102.30              -19.7  Mercado Livre
P-6Y2SPRR4G4           85.91     62.00          -27.9          71.50              -16.8  Mercado Livre
...

✅ Análise salva em: resultado_analise_precos.xlsx
```

---

## 📊 PASSO 3: ENTENDER OS RESULTADOS

### 3.1 Arquivo de Saída

Um novo arquivo será criado: `resultado_analise_precos.xlsx`

**Sheet 1 - "Comparativo Preços":**
```
Coluna A: ID_Produto              → ID interno da Sportbay
Coluna B: Nome_Produto            → Nome do produto
Coluna C: Preco_Sportbay          → Seu preço
Coluna D: Preco_ML                → Melhor preço no Mercado Livre
Coluna E: Diferenca_ML_%          → Variação % (+ = mais caro, - = mais barato)
Coluna F: Link_ML                 → Link direto do produto
Coluna G: Preco_Shopee            → Melhor preço na Shopee
Coluna H: Diferenca_Shopee_%      → Variação %
Coluna I: Link_Shopee             → Link direto do produto
Coluna J: Melhor_Marketplace      → Qual tem melhor preço
```

**Sheet 2 - "Resumo Estatístico":**
```
Métrica                          Valor
Total de Produtos                9
Matches ML                        7
Matches Shopee                    6
Taxa Match ML (%)                77.78
Taxa Match Shopee (%)            66.67
Preço Médio Sportbay             R$ 158.41
Preço Médio ML                   R$ 128.75
Preço Médio Shopee               R$ 135.20
Diferença Média ML (%)           -18.7%
Diferença Média Shopee (%)       -14.2%
Produtos Mais Caros no ML        1
Produtos Mais Caros na Shopee    2
```

### 3.2 Como Interpretar os Dados

**Diferença % Positiva (+):**
```
Exemplo: Diferenca_ML_% = +15.0
Significado: O Mercado Livre vende 15% MAIS CARO que Sportbay
Ação: Você é competitivo! ✓
```

**Diferença % Negativa (-):**
```
Exemplo: Diferenca_ML_% = -25.5
Significado: O Mercado Livre vende 25.5% MAIS BARATO que Sportbay
Ação: Considere revisar seu preço ⚠️
```

**Null (Vazio):**
```
Significado: Não encontrou o produto no marketplace
Causa: Nome muito diferente ou não está cadastrado
Ação: Verificar manualmente
```

---

## 🔄 PASSO 4: ATUALIZAR ANÁLISE PERIODICAMENTE

### 4.1 Roda Diária/Semanal

Modifique o script para automatizar:

```python
"""
Script de monitoramento periódico
"""

import asyncio
import pandas as pd
from datetime import datetime
from pricing_intelligence_pro_tork import main

# Arquivo de histórico
ARQUIVO_HISTORICO = 'historico_precos.xlsx'

async def coletar_preco():
    # Carregar produtos
    df_entrada = pd.read_excel('analise_de_preços.xlsx')
    
    # Analisar
    df_resultado = await main(df_entrada=df_entrada)
    
    # Adicionar timestamp
    df_resultado['data_coleta'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Salvar resultado
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    arquivo_saida = f'resultado_{timestamp}.xlsx'
    
    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, sheet_name='Comparativo', index=False)
    
    return df_resultado

# Executar
if __name__ == '__main__':
    print(f"🕐 Coleta iniciada às {datetime.now().strftime('%H:%M:%S')}")
    df = asyncio.run(coletar_preco())
    print(f"✅ Análise concluída à {datetime.now().strftime('%H:%M:%S')}")
```

### 4.2 Agendar Execução Automática

**Windows (Task Scheduler):**
```
1. Abrir: Agendador de Tarefas
2. Criar tarefa básica
3. Disparador: Diariamente às 18:00
4. Ação: Executar "python" com argumento "executar_analise.py"
```

**Mac/Linux (Cron):**
```bash
# Editar crontab
crontab -e

# Adicionar linha (executar toda segunda-feira às 09:00):
0 9 * * 1 cd /caminho/do/projeto && python3 executar_analise.py
```

---

## 🎓 PASSO 5: ANÁLISES AVANÇADAS (OPCIONAL)

### 5.1 Filtrar Apenas Produtos Não Competitivos

```python
import pandas as pd

# Carregar resultado anterior
df = pd.read_excel('resultado_analise_precos.xlsx', sheet_name='Comparativo Preços')

# Produtos onde Sportbay é significativamente mais caro
nao_competitivos = df[
    ((df['Diferenca_ML_%'] < -20) | (df['Diferenca_Shopee_%'] < -20)) &
    (df['Preco_Sportbay'].notna())
]

print("🔴 PRODUTOS NÃO COMPETITIVOS (20%+ mais caro):\n")
for _, row in nao_competitivos.iterrows():
    print(f"  {row['Nome_Produto']}")
    print(f"    Sportbay: R${row['Preco_Sportbay']:.2f}")
    print(f"    ML:       R${row['Preco_ML']:.2f} ({row['Diferenca_ML_%']:+.1f}%)")
    print(f"    Shopee:   R${row['Preco_Shopee']:.2f} ({row['Diferenca_Shopee_%']:+.1f}%)")
    print()
```

### 5.2 Comparar com Semana Anterior

```python
import pandas as pd

# Carregar análise anterior
df_anterior = pd.read_excel('resultado_20240115_090000.xlsx')
df_nova = pd.read_excel('resultado_analise_precos.xlsx', sheet_name='Comparativo Preços')

# Mesclar dados
comparacao = df_nova.merge(
    df_anterior[['ID_Produto', 'Preco_ML', 'Preco_Shopee']],
    on='ID_Produto',
    suffixes=('_novo', '_anterior')
)

# Calcular variações
comparacao['Variacao_ML'] = (
    comparacao['Preco_ML_novo'] - comparacao['Preco_ML_anterior']
)

print("📊 VARIAÇÃO DE PREÇOS (Semana Anterior → Agora):\n")
for _, row in comparacao.iterrows():
    var_ml = row['Variacao_ML']
    sinal = "📈" if var_ml > 0 else "📉"
    print(f"{sinal} {row['Nome_Produto']}")
    print(f"   ML: R${row['Preco_ML_anterior']:.2f} → R${row['Preco_ML_novo']:.2f} ({var_ml:+.2f})")
```

### 5.3 Exportar Apenas Competitivos para Análise

```python
import pandas as pd

df = pd.read_excel('resultado_analise_precos.xlsx', sheet_name='Comparativo Preços')

# Filtrar: Sportbay é mais barato OU preço similar
competitivos = df[
    (df['Diferenca_ML_%'] >= -10) &
    (df['Diferenca_Shopee_%'] >= -10)
]

# Salvar em planilha simples
competitivos[[
    'ID_Produto',
    'Nome_Produto',
    'Preco_Sportbay',
    'Melhor_Marketplace',
    'Diferenca_ML_%',
    'Diferenca_Shopee_%'
]].to_excel('produtos_competitivos.xlsx', index=False)

print(f"✅ {len(competitivos)}/{len(df)} produtos são competitivos")
```

---

## 🐛 SOLUÇÃO DE PROBLEMAS

### ❌ Erro: "ModuleNotFoundError"

```bash
# Solução:
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### ❌ Erro: "No module named 'playwright'"

```bash
# Solução:
playwright install chromium
```

### ❌ Script muito lento

```python
# Se tem >100 produtos, dividir em lotes:

produtos_total = pd.read_excel('analise_de_preços.xlsx')
tamanho_lote = 20

resultados = []
for i in range(0, len(produtos_total), tamanho_lote):
    lote = produtos_total[i:i+tamanho_lote]
    print(f"Processando lote {i//tamanho_lote + 1}...")
    df = asyncio.run(main(df_entrada=lote))
    resultados.append(df)

df_final = pd.concat(resultados, ignore_index=True)
```

### ❌ "ConnectionError" ou timeout

Tentar novamente - geralmente é problema de conexão temporária.

### ❌ Muitos "Não encontrados" (NULL)

Verificar nomes dos produtos - muito específicos ou diferentes do mercado:

```python
# Simplificar nomes
df['nome_produto'] = df['nome_produto'].str.replace('PRO TORK', 'Pro Tork')
df['nome_produto'] = df['nome_produto'].str.replace('NEW', '')
```

---

## 📈 MÉTRICAS ESPERADAS

Para seus 9 produtos Pro Tork:

```
Taxa de Match Esperada:
  - Mercado Livre: 70-90% (API confiável)
  - Shopee: 60-80% (web scraping, mais variável)

Tempo de Processamento:
  - 9 produtos: ~20-40 segundos

Diferença de Preço Típica (Pro Tork):
  - Capacetes: -15% a +25%
  - Acessórios: -20% a +30%
  - Média geral: -12% a +5%
```

---

## ✅ CHECKLIST FINAL

- [ ] Python 3.8+ instalado (`python --version`)
- [ ] Dependências instaladas (`pip list` mostra pandas, requests, playwright, rapidfuzz)
- [ ] Playwright instalado (`playwright install chromium`)
- [ ] Planilha na pasta do projeto (`analise_de_preços.xlsx`)
- [ ] Script de execução criado (`executar_analise.py`)
- [ ] Teste rápido (`python setup_helper.py test`)
- [ ] Primeira execução (`python executar_analise.py`)
- [ ] Resultado gerado (`resultado_analise_precos.xlsx` criado)

---

## 🎯 PRÓXIMOS PASSOS

1. **Executar análise inicial** → Entender padrão de preços
2. **Identificar produtos não-competitivos** → Possíveis ações de precificação
3. **Agendar coleta periódica** → Monitoramento contínuo
4. **Integrar com BI** → Dashboard de preços (Tableau, Power BI, etc)
5. **Exportar para API interna** → Atualizar preços automaticamente (futura integração)

---

**Dúvidas? Referir-se ao README.md ou aos comentários no código Python!**
