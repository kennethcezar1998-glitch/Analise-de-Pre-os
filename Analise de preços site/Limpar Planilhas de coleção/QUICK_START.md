╔════════════════════════════════════════════════════════════════════════════╗
║                          🚀 QUICK START (5 MIN)                           ║
║           Pricing Intelligence - Pro Tork Marketplace Monitor              ║
╚════════════════════════════════════════════════════════════════════════════╝

## 📋 O QUE VOCÊ TEM

✅ 9 produtos Pro Tork em sua planilha
✅ Todos os scripts prontos para usar
✅ Documentação completa

## ⚡ COMEÇAR EM 5 MINUTOS

### 1️⃣ INSTALAR (COPIAR E COLAR)

```bash
# Windows
pip install -r requirements.txt
python -m playwright install chromium

# Mac/Linux
pip3 install -r requirements.txt
python3 -m playwright install chromium
```

**Tempo: ~5 minutos** ⏱️

Procure por:
```
Successfully installed pandas-...
✓ Chromium browser downloaded
```

### 2️⃣ EXECUTAR ANÁLISE (COPIAR E COLAR)

```bash
# Windows
python executar_analise.py

# Mac/Linux
python3 executar_analise.py
```

**Tempo: ~2 minutos** ⏱️

Você verá:
```
🚀 PRICING INTELLIGENCE SYSTEM

📂 CARREGANDO PLANILHA
✓ Planilha carregada: analise_de_preços.xlsx
✓ Produtos encontrados: 9

⏳ EXECUTANDO ANÁLISE
Buscando preços em Mercado Livre e Shopee...

📊 RESUMO
   Total: 9
   Mercado Livre: 7 encontrados
   Shopee: 6 encontrados

✅ ANÁLISE COMPLETA
📊 Arquivo salvo: resultado_analise_precos.xlsx
```

### 3️⃣ VERIFICAR RESULTADOS

Abra o arquivo gerado: **`resultado_analise_precos.xlsx`**

Você verá:
- **Sheet 1**: Comparativo detalhado com links dos produtos
- **Sheet 2**: Resumo estatístico
- **Sheet 3**: Produtos não competitivos

---

## 🎯 ENTENDER OS RESULTADOS (2 MIN)

### Exemplo de Resultado:

```
Seu produto:      CAPACETE PRO TORK NEW LIBERTY 3
Preço Sportbay:   R$ 127.48

Mercado Livre:    R$ 95.00  (-25.5%) ← 25% MAIS BARATO que você
Shopee:           R$ 102.30 (-19.7%) ← 19% MAIS BARATO que você

Melhor:           Mercado Livre
```

### Interpretação:

✅ **Seu preço é competitivo?**
- ✓ Sim = Diferença < -15% (Você é mais barato ou similar)
- ✗ Não = Diferença < -20% (Precisa revisar preço)

---

## 📊 ARQUIVOS GERADOS

```
resultado_analise_precos.xlsx
│
├─ Sheet: Comparativo Preços
│  └─ ID, Nome, Preço, Links, Diferenças
│
├─ Sheet: Resumo Executivo
│  └─ Estatísticas gerais
│
└─ Sheet: Produtos Não Competitivos
   └─ Produtos onde você é significativamente mais caro
```

---

## 🐛 PROBLEMAS COMUNS

### "ModuleNotFoundError"
```bash
pip install -r requirements.txt
```

### "Chromium not found"
```bash
playwright install chromium
```

### "File not found"
- Coloque `analise_de_preços.xlsx` na mesma pasta do script
- Ou renomeie seu arquivo para exatamente esse nome

### Script muito lento
- Normal: 2-3 minutos para 9 produtos
- Shopee pode demorar: está carregando dados dinâmicos

### Muitos "não encontrados"
- Alguns produtos podem estar com nomes muito específicos
- Resultado: vai aparecer como "N/A" (não encontrado)
- Solução: verificar manualmente no marketplace

---

## 🔄 PRÓXIMOS PASSOS

### Próxima semana:
```bash
python executar_analise.py
```
Arquivo novo será criado: `resultado_analise_precos.xlsx`

### Acompanhamento mensal:
Abra os resultados e procure por:
- 🔴 Produtos onde você é muito mais caro
- 🟢 Produtos onde você é competitivo
- 📈 Tendência de variação de preços

---

## 📞 PRECISA DE AJUDA?

### Verificar logs:
```bash
cat pricing_intelligence.log  # Mac/Linux
type pricing_intelligence.log # Windows
```

### Testar conexão:
```bash
python setup_helper.py test
```

### Ver demo de teste:
```bash
python setup_helper.py demo
```

---

## 🎓 PRÓXIMAS ANÁLISES AVANÇADAS

Quando se sentir confortável, explore:

✓ Análises de histórico de preços
✓ Comparação com semana anterior
✓ Filtros customizados
✓ Integração com seu sistema interno
✓ Alertas automáticos para preços fora da faixa

---

## ✅ CHECKLIST

- [ ] Python instalado (`python --version`)
- [ ] requirements.txt instalado (`pip list` mostra pandas, requests, playwright)
- [ ] Playwright instalado (chromium baixado)
- [ ] `analise_de_preços.xlsx` na pasta
- [ ] Primeira execução (`python executar_analise.py`)
- [ ] Resultado gerado (`resultado_analise_precos.xlsx`)
- [ ] Aberto e analisado o resultado

---

**Tudo pronto? Execute agora:**

```bash
python executar_analise.py
```

**Bom monitoramento! 🚀**
