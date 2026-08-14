# Bot de Enriquecimento de SKU - Sportbay

Automatiza o cadastro de Enriquecimento de SKU no backoffice
(`OPERAÇÃO > Enriquecimento de SKU > Criar`), lendo os dados da planilha
`Tags para enriquecimento.xlsx` (aba `Tags`) na pasta acima.

## Como funciona

- Coluna A -> Código SKU da Sportbay
- Coluna B -> Código SKU do seller
- Coluna C -> Curva ABC
- Coluna D -> Categoria (filtro "Categorias predefinidas", match exato)
- Coluna E -> Grupo (filtro "Filtrar por grupo", usa a aba `Listas` para achar o código)
- Colunas F e G -> Facet / SPEC, usados para achar a tag certa no filtro "Tag global"
  (aceita pequenas divergências de texto, mas se não achar nada parecido o suficiente
  a tag é **pulada** e listada no final do log, para revisão manual - o bot nunca
  clica em uma tag "chutada").

Linhas consecutivas com o mesmo par (SKU Sportbay, SKU seller) são tratadas como
**um único cadastro**: o bot preenche SKU/Seller/Curva/Categoria uma vez, adiciona
uma Tag Global para cada linha do grupo, e só clica em "Salvar Enriquecimento"
depois da última linha do grupo.

## Instalação (uma vez)

```bash
cd "bot"
pip install -r requirements.txt
python -m playwright install chromium
```

## Configuração

As credenciais já estão no arquivo `.env` desta pasta (`SPORTBAY_USER` /
`SPORTBAY_PASS`). Troque a senha lá se você mudar a senha da conta.

**Atenção:** a senha fica em texto puro nesse arquivo. Não compartilhe essa pasta
nem suba `.env` para nenhum repositório.

## Rodando

Primeiro teste, sempre com `--dry-run` e `--limit 1` para conferir visualmente se
tudo está sendo preenchido certo (o navegador abre visível por padrão):

```bash
python enrich_bot.py --dry-run --limit 1
```

Se estiver tudo certo, rode de verdade (ainda para 1 grupo, agora salvando):

```bash
python enrich_bot.py --limit 1
```

Depois de validado, rode tudo:

```bash
python enrich_bot.py
```

### Opções

- `--dry-run` - preenche o formulário inteiro mas nunca clica em Salvar.
- `--headless` - roda sem abrir a janela do navegador (mais rápido, sem supervisão visual).
- `--limit N` - processa só os N primeiros grupos de SKU.
- `--start-from N` - começa a partir do grupo N (1-based), útil para retomar depois de uma falha.

## Logs

Cada execução grava em `bot/logs/run.log`. Erros e avisos (ex.: tag sem
correspondência, botão Salvar que não habilitou) tiram um screenshot em
`bot/logs/` para facilitar o diagnóstico. No fim da execução o log lista todas
as tags puladas por falta de correspondência confiável, com o número da linha
na planilha original.

## Limitações conhecidas / pontos para ajustar após o primeiro teste

O robô foi montado a partir de capturas de HTML do site (arquivos `HTML*.txt`
na pasta acima), mas os pop-ups de "Categorias predefinidas" e "Tag global"
não foram capturados abertos - os seletores usados (`get_by_label` +
`role="option"` dentro de `role="listbox"`) são o padrão do Material UI e devem
funcionar, mas **o primeiro `--dry-run --limit 1` é essencial** para confirmar
isso antes de rodar em lote. Se algum passo falhar, o log e o screenshot en
`bot/logs/` mostram exatamente onde parou.
