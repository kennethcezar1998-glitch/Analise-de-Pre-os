# Briefing — Correção do bug de parágrafos vazios no editor SEO (Sportbay)

## Contexto

No admin do marketplace Sportbay, as telas de edição de **coleções, produtos e departamentos** têm o mesmo componente de texto rico (Quill) para o conteúdo SEO da página — o wrapper tem `id="seo"` e o conteúdo editável fica em `#seo .ql-editor`.

## O problema

A cada "salvar"/"atualizar", o editor reinsere **dois parágrafos vazios** (`<p><br></p><p><br></p>`) imediatamente **antes de cada heading que vem depois de um parágrafo de texto**. Em produção isso vira um vão enorme entre o texto e o título seguinte.

Diagnóstico feito comparando dois dumps de HTML da **mesma página** (um com o bug, outro depois da limpeza manual):

- diferença = exatamente **22 blocos `<p><br></p>`**, sempre em pares;
- removendo todos eles, o HTML com bug fica **byte a byte idêntico** ao HTML correto;
- heading colado em heading **não** recebe o lixo (ex.: `<h2>Perguntas frequentes</h2>` seguido do primeiro `<h3>` do FAQ está limpo nos dois arquivos).

## O que foi pedido

Que o bot, **antes** de remover os produtos em "Produtos da categoria" ou adicionar em "Seleção Manual de Produtos", entre no componente de texto e o normalize — dando "Backspace" até o conteúdo voltar à estrutura correta.

Decisões de escopo tomadas:

1. **Momento:** logo depois de abrir a coleção e **antes** de mexer nos produtos. O `salvar()` do fim do fluxo é o mesmo para tudo, então a correção do SEO pega carona nele — sem clique de salvar extra e sem reload (que reintroduziria o bug).
2. **Regra:** remover **somente** sequências de parágrafos vazios que antecedem imediatamente um heading (h1–h6). Parágrafo vazio em qualquer outra posição (entre dois parágrafos, no fim do texto, antes de lista) é **preservado** — pode ser espaçamento intencional de quem escreveu o conteúdo.
3. **Reutilização:** implementado como módulo separado, para os bots de produtos e departamentos importarem depois.

## O que foi implementado

### Arquivo novo: `limpador_seo.py`

Função pública:

```python
limpar_paragrafos_vazios_seo(driver, seletores=None, log=None)
# -> {'editor_encontrado', 'removidos', 'restantes', 'metodo', 'texto_preservado'}
```

Nunca levanta exceção — qualquer falha é logada e devolvida no dict, para a atualização de produtos não travar por causa do SEO.

**Por que não mexer no `innerHTML`:** o Quill mantém um modelo interno (Delta) e o React lê desse modelo, não do DOM. Reescrever o `innerHTML` na marra não dispara `text-change`, então o "Salvar" mandaria o conteúdo **velho** para o backend. Daí os dois caminhos, nessa ordem:

1. **API do Quill** — pega a instância (`container.__quill` / `Quill.find(container)`) e chama `quill.deleteText(idx, 1, 'user')` em cada linha vazia alvo, **de trás para frente** (apagar a última primeiro mantém válidos os índices das anteriores). O source `'user'` faz o React-Quill propagar o `onChange`. Guarda de segurança: só apaga blot com `length() === 1` (parágrafo vazio é só o `\n`).
2. **Fallback por teclado** — se a instância não estiver exposta (Quill 2 guarda as instâncias num WeakMap, sem `__quill`), posiciona o cursor no parágrafo vazio via `Range`/`Selection` e envia uma tecla **real** pelo Selenium (`ActionChains`), uma por vez, relendo o DOM a cada volta (o Quill reconstrói os nós depois de cada edição → referência guardada daria `StaleElementReference`). O `MutationObserver` do Quill sincroniza o modelo sozinho.

**Detalhe importante — DELETE, não BACKSPACE:** no Quill, o caractere `\n` é quem carrega a formatação do bloco que ele termina. Com o cursor no início do heading, um BACKSPACE apagaria o `\n` do parágrafo vazio de cima e o texto do título subiria para dentro de um `<p>`, **perdendo a tag de heading**. Com o cursor dentro do parágrafo vazio, o DELETE apaga o `\n` do próprio vazio: o `\n` do heading sobrevive e a formatação do título é mantida.

**Conferência de integridade:** o texto visível é capturado antes e depois (ignorando espaços, NBSP e zero-width) e comparado. Se algo além de linha vazia sumir, vira ERRO no log e `texto_preservado=False`.

Constantes ajustáveis no topo do módulo: `SELETORES_EDITOR`, `MAX_ITERACOES_TECLADO` (300), `PAUSA_ENTRE_TECLAS` (0.05s).

### Alterações em `robo_colecoes_sportbay.py`

Três mudanças pontuais, nada mais foi tocado:

1. `from limpador_seo import limpar_paragrafos_vazios_seo` (os dois arquivos precisam ficar na mesma pasta ou o módulo no `PYTHONPATH`);
2. nova função `normalizar_seo(driver)`, logo acima de `salvar()`;
3. chamada no `main()`, entre `abrir_colecao()` e `remover_todos_produtos_atuais()`.

Comportamento de `normalizar_seo()`: se `texto_preservado` for `False`, **aborta o fluxo** com `RuntimeError`. Como nenhum produto foi tocado ainda, abortar não deixa a coleção pela metade — e evita que o `salvar()` do fim persista conteúdo danificado. Editor ausente ou vazio apenas loga e segue.

## Testes feitos

Lógica JS testada com **jsdom**, usando o mesmo código que roda no navegador (extraído do próprio módulo, sem cópia manual) contra os dois HTMLs reais da página:

- 52 blocos, 22 alvos detectados, 22 removidos, 0 restantes;
- texto preservado: OK;
- HTML resultante **idêntico** ao HTML alvo.

Casos de borda validados:

| Caso | Resultado |
|---|---|
| vazio entre dois parágrafos | preservado |
| vazio no fim do texto | preservado |
| vazio antes de lista (`<ul>`) | preservado |
| vazio(s) antes de heading | removido(s) |
| `<p>&nbsp;</p>` antes de heading | removido |
| heading colado em heading | intacto |
| vazio no início, antes do `<h1>` | removido |

## Bug corrigido — "Editor de SEO não encontrado" (rodada de 18/08)

Na primeira rodada real o log deu:

```
15:54:03,758 - INFO - ✓ Coleção aberta para edição
15:54:03,763 - INFO - Editor de SEO não encontrado nesta tela — nada a limpar
```

**5 milissegundos** entre as duas linhas — o módulo consultava o DOM na hora e
desistia. Os dumps de HTML da mesma página mostram que o componente está lá e o
seletor está certo: `<div id="seo" class="quill">` → `.ql-container` →
`.ql-editor`, e o bloco fica no meio do formulário, antes de "Produtos da
categoria", sem aba nem accordion escondendo nada.

**Causa:** corrida com o React. `abrir_colecao()` dá a página por aberta quando
o título "Editar categoria/coleção" aparece, mas o formulário continua montando
depois disso, e o SEO fica no fim dele. Havia ainda um segundo atraso: o editor
monta vazio e só recebe o texto quando a resposta do backend chega.

**Correção (só em `limpador_seo.py` — o fluxo de remover/adicionar/salvar não
foi tocado):** a inspeção agora espera em três estágios antes de concluir
qualquer coisa —

1. o editor aparecer no DOM (`TIMEOUT_EDITOR`, 20s);
2. o texto chegar (`TIMEOUT_CONTEUDO`, 10s) — editor vazio depois disso é
   coleção sem SEO mesmo, e o módulo só loga e segue;
3. o texto se repetir em duas leituras seguidas, para a limpeza não começar no
   meio do preenchimento (o que faria a conferência de integridade comparar
   textos de momentos diferentes e abortar o fluxo à toa).

As mensagens de log passaram a dizer quanto se esperou ("não encontrado após
20s", "continuou vazio após 10s"), para a próxima rodada distinguir ausência real
de espera curta.

Outros dois ajustes de robustez no mesmo arquivo:

- **Fallback por teclado com medidor de progresso:** o JS que posiciona o cursor
  agora devolve também quantos alvos restam. O laço credita a remoção pela queda
  da contagem (antes contava iteração, não remoção) e para depois de
  `MAX_TENTATIVAS_SEM_PROGRESSO` (3) varreduras sem a contagem cair, em vez de
  girar até as 300 iterações mandando DELETE no vazio.
- **Conferência final com segunda chance:** dá um respiro para o React
  re-renderizar e, se o texto vier diferente, relê antes de acusar perda de
  conteúdo — um falso positivo ali aborta a rodada inteira.

Validação: a regra de detecção foi reconferida contra os dois HTMLs reais da
pasta `Bug do conteúdo SEO` (52 blocos de topo, 22 parágrafos vazios, **22
alvos**, e o HTML com bug menos os 22 `<p><br></p>` bate byte a byte com o
arquivo correto). A lógica de espera e o fallback foram exercitados com um driver
falso em 7 cenários: editor que monta tarde, editor que monta vazio e recebe o
texto depois, tela sem SEO, coleção sem texto, SEO já limpo, ausência da
instância do Quill (cai no teclado) e DELETE sem efeito (para na guarda de
progresso).

## Rodada de validação em produção (sem salvar)

Teste real na coleção 'Mais Potência', sem clicar em Salvar e sem tocar nos
produtos:

```
16:12:35,526 - ✓ Coleção aberta para edição
16:12:35,531 - [TESTE] Diagnóstico IMEDIATO (como era antes da correção): {'editor': False}
16:12:36,823 - Editor de SEO: 52 blocos, 22 parágrafo(s) vazio(s), 22 antes de heading
16:12:37,332 - ✓ SEO normalizado: 22 parágrafo(s) vazio(s) removido(s) via quill
```

A consulta instantânea reproduziu o bug (`editor: False`) e o editor apareceu
~1,3 s depois — exatamente a corrida diagnosticada. Do "coleção aberta" ao SEO
limpo foram **1,8 s**.

O `innerHTML` do editor depois da limpeza tem 5375 caracteres, zero
`<p><br></p>`, e é **byte a byte idêntico** ao dump `Como o conteúdo SEO deve
ficar.txt`.

**`metodo` = `quill`** — a instância está exposta na página, então o caminho
rápido (`deleteText(..., 'user')`) resolve tudo e o fallback por teclado não
chegou a rodar. Não há motivo, por ora, para otimizar os DELETEs em lote.

## Pontos em aberto / próximos passos
- **Origem do bug ainda não confirmada:** se o lixo for injetado pela própria rotina de *save* do admin (e não na renderização ao abrir a página), limpar antes dos produtos não basta — o clique em "Salvar" reintroduziria tudo. Nesse caso, mover (ou duplicar) a chamada para imediatamente antes do `salvar()`. A primeira rodada em produção resolve essa dúvida.
- **Portar para os bots de produtos e departamentos:** basta importar a mesma função; os seletores já cobrem telas sem o `id="seo"`.
