# Bot de EDIÇÃO de Enriquecimento de SKU - Sportbay

Automatiza a **edição** de registros já cadastrados em
`OPERAÇÃO > Enriquecimento de SKU > Lista`, lendo os dados da planilha
`Tags para enriquecimento - edição.xlsx` (aba `Tags`) na pasta acima.

Este bot é irmão do bot de **criação** (pasta `bot` no nível acima desta,
usado em `OPERAÇÃO > Enriquecimento de SKU > Criar`). A lógica de preencher
"Filtrar por grupo" + "Tag global" é idêntica. As diferenças são:

- Em vez de ir para a tela de **Criar**, o bot vai para a **Lista**, digita o
  SKU no campo "Buscar por SKU" e abre o registro pelo botão "Editar".
- Ele **nunca** preenche "Código SKU da Sportbay" / "Código SKU do seller"
  (não são editáveis na tela de edição).
- Ele **nunca** altera "Curva ABC" nem "Categorias predefinidas" - só
  adiciona novas Tags Globais.
- Se a busca por SKU retornar mais de um resultado, o bot usa sempre o
  **primeiro** da lista (fica registrado no log quando isso acontece, para
  conferência).

## Como funciona

- Coluna A -> Código SKU da Sportbay (usado para buscar o registro na Lista)
- Coluna B -> Código SKU do seller (só para conferência/agrupamento, não é
  digitado em lugar nenhum)
- Coluna C (Curva ABC) e D (Categorias) -> **ignoradas** por este bot
- Coluna E -> Grupo (filtro "Filtrar por grupo", usa a aba `Listas` para achar o código)
- Colunas F e G -> Facet / SPEC, usados para achar a tag certa no filtro "Tag global"
  (aceita pequenas divergências de texto, mas se não achar nada parecido o suficiente
  a tag é **pulada** e listada no final do log, para revisão manual - o bot nunca
  clica em uma tag "chutada").

Linhas consecutivas com o mesmo par (SKU Sportbay, SKU seller) são tratadas como
**um único registro para editar**: o bot abre a tela de edição dele uma vez,
adiciona uma Tag Global para cada linha do grupo, e só clica em "Atualizar
Enriquecimento" depois da última linha do grupo.

Se um SKU da planilha não for encontrado na Lista, o grupo é pulado e listado
no final do log em "SKU(s) não encontrados".

## Instalação (uma vez)

```bash
cd "bot"
pip install -r requirements.txt
python -m playwright install chromium
```

## Configuração

O arquivo `.env` desta pasta já foi copiado do bot de criação (mesmas
credenciais `SPORTBAY_USER` / `SPORTBAY_PASS`, mesmo login). Troque a senha
lá se você mudar a senha da conta.

**Atenção:** a senha fica em texto puro nesse arquivo. Não compartilhe essa
pasta nem suba `.env` para nenhum repositório.

## Rodando

Primeiro teste, sempre com `--dry-run` e `--limit 1` para conferir visualmente se
tudo está sendo preenchido certo (o navegador abre visível por padrão):

```bash
python edit_bot.py --dry-run --limit 1
```

Se estiver tudo certo, rode de verdade (ainda para 1 grupo, agora salvando):

```bash
python edit_bot.py --limit 1
```

Depois de validado, rode tudo:

```bash
python edit_bot.py
```

### Opções

- `--dry-run` - preenche o formulário inteiro mas nunca clica em Salvar.
- `--headless` - roda sem abrir a janela do navegador (mais rápido, sem supervisão visual).
- `--limit N` - processa só os N primeiros registros.
- `--start-from N` - começa a partir do registro N (1-based), útil para retomar depois de uma falha.

## Logs

Cada execução grava em `bot/logs/run.log`. Erros e avisos (ex.: SKU não
encontrado na Lista, tag sem correspondência, botão Atualizar Enriquecimento
que não habilitou) tiram um screenshot em `bot/logs/` para facilitar o
diagnóstico. No fim da execução o log lista todos os SKUs não encontrados e
todas as tags puladas por falta de correspondência confiável, com o número
da linha na planilha original.

## Testado em produção (dry-run)

A busca "Buscar por SKU" da Lista **não é um filtro exato** - digitar um SKU
pode devolver várias linhas, inclusive de outros SKUs (confirmado num teste
real: buscar "110731" trouxe 3 resultados, um deles de outro SKU). Por isso
o bot:

1. Só abre a linha da tabela cujo SKU Sportbay **e** SKU seller batem
   exatamente com os da planilha (`find_edit_url`), nunca "o primeiro".
2. Depois de abrir a tela de edição, confere de novo os valores exibidos nos
   campos bloqueados de SKU antes de mexer em qualquer coisa
   (`confirm_edit_page_sku`). Se não bater, aborta aquele grupo e tira
   screenshot em vez de arriscar editar o registro errado.

O botão de salvar na tela de edição chama **"Atualizar Enriquecimento"**
(diferente do "Salvar Enriquecimento" da tela de Criar) - já ajustado no
código. O layout dos campos "Filtrar por grupo" / "Tag global" é o mesmo da
tela de Criar, confirmado em teste real (`--dry-run --limit 1`).

## Limitações conhecidas

Se algum passo falhar, o log e o screenshot em `bot/logs/`
mostram exatamente onde parou.
