"""
Limpeza do componente de texto SEO (Quill) — Sportbay Marketplace

Módulo reutilizável: serve para coleções, produtos e departamentos, já que o
componente de SEO é o mesmo em todas as telas de edição do admin.

--- O BUG ---
A cada "salvar"/"atualizar", o editor reinsere dois parágrafos vazios
(`<p><br></p><p><br></p>`) imediatamente ANTES de cada heading que vem depois de
um parágrafo de texto. Em produção isso vira um vão enorme entre o texto e o
título seguinte. Heading colado em heading (ex.: `<h2>Perguntas frequentes</h2>`
seguido do primeiro `<h3>` do FAQ) não recebe o lixo.

Comparando o HTML com bug e o HTML correto da mesma página, a diferença é
exatamente 22 blocos `<p><br></p>`: removendo todos eles, o HTML com bug fica
byte a byte idêntico ao HTML desejado.

--- A REGRA APLICADA AQUI ---
Remove SOMENTE sequências de parágrafos vazios que antecedem imediatamente um
heading (h1–h6). Parágrafo vazio em qualquer outra posição (entre dois
parágrafos, no fim do texto) é preservado — pode ser espaçamento intencional de
quem escreveu o conteúdo.

--- POR QUE NÃO MEXER NO innerHTML ---
O Quill mantém um modelo interno (Delta) e o React lê desse modelo, não do DOM.
Reescrever o `innerHTML` na marra não dispara `text-change`, então o "Salvar"
mandaria o conteúdo VELHO para o backend (ou um estado inconsistente). Por isso
há dois caminhos, nessa ordem:

  1. API do Quill (`quill.deleteText(..., 'user')`) — propaga o evento
     normalmente e o React atualiza o estado. É o caminho rápido e preferido.
  2. Fallback por teclado — posiciona o cursor no parágrafo vazio e manda um
     DELETE de verdade. O navegador edita o contenteditable e o MutationObserver
     do Quill sincroniza o modelo sozinho. É o "Backspace na mão", só que com
     DELETE (ver nota abaixo).

NOTA — por que DELETE e não BACKSPACE: no Quill, o caractere `\n` é quem carrega
a formatação do bloco que ele termina. Com o cursor no início do heading, um
BACKSPACE apagaria o `\n` do parágrafo vazio de cima e o texto do heading subiria
para dentro de um `<p>`, PERDENDO a tag de heading. Com o cursor no parágrafo
vazio, o DELETE apaga o `\n` do próprio parágrafo vazio: o `\n` do heading
sobrevive e a formatação do título é mantida. É a diferença entre limpar o
espaçamento e destruir o H2 da página.

--- POR QUE ESPERAR O EDITOR ---
A tela de edição é React: o título "Editar categoria/coleção" (que o bot usa
para dar a página por aberta) aparece ANTES de o formulário terminar de montar,
e o componente de SEO fica lá embaixo, montado depois — e ainda por cima monta
vazio, recebendo o texto só quando a resposta do backend chega. Consultar o DOM
na hora dava `#seo .ql-editor` inexistente e o módulo concluía "editor não
encontrado" milissegundos depois de abrir a coleção. Por isso a inspeção agora
espera, em duas etapas: o componente aparecer (TIMEOUT_EDITOR) e o conteúdo
chegar e parar de mudar (TIMEOUT_CONTEUDO), antes de decidir que não há editor
ou que ele está vazio.

--- SEGURANÇA ---
- Se o editor não existir na página (tela sem componente de SEO) ou estiver
  vazio, a função apenas loga e devolve status — nunca levanta exceção, para não
  derrubar o fluxo principal do bot.
- O texto visível é capturado antes e depois (ignorando espaços em branco) e
  comparado: se algo além de linha vazia tiver sumido, isso vira ERRO no log e
  `texto_preservado=False` no retorno, para quem chamou decidir se salva ou não.
"""
import time
import logging

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

logger = logging.getLogger(__name__)

# O wrapper do componente tem id="seo" no admin; o segundo seletor cobre telas
# onde o id não esteja presente, pegando o primeiro editor Quill da página.
SELETORES_EDITOR = ["#seo .ql-editor", ".ql-editor"]

# Teto de segurança do fallback por teclado: cada iteração remove 1 parágrafo.
# Uma página cheia tem ~22 alvos; 300 é folga suficiente sem risco de loop
# infinito caso alguma tecla não produza efeito.
MAX_ITERACOES_TECLADO = 300

# Pausa entre cada DELETE no fallback — dá tempo do Quill sincronizar o modelo
# via MutationObserver antes de a próxima varredura ler o DOM.
PAUSA_ENTRE_TECLAS = 0.05

# Quantas varreduras seguidas sem o número de alvos cair antes de desistir do
# fallback por teclado. Sem isso, um DELETE que não surte efeito faria o laço
# girar até MAX_ITERACOES_TECLADO sem remover nada.
MAX_TENTATIVAS_SEM_PROGRESSO = 3

# Espera pelo componente de SEO montar na tela de edição (React monta o
# formulário em partes; o SEO fica no fim dele).
TIMEOUT_EDITOR = 20

# Depois de montado, o editor ainda aparece vazio até o texto chegar do
# backend. Este é o teto para esperar o conteúdo — estourou e continuou vazio,
# a coleção realmente não tem SEO escrito.
TIMEOUT_CONTEUDO = 10

# Intervalo entre as varreduras das esperas acima.
INTERVALO_POLL = 0.25


# Trecho JS compartilhado: define `acharEditor`, `ehVazio` e `alvosDom`.
# `ehVazio` ignora espaço comum, NBSP e zero-width — o Quill usa `<br>` para
# linha vazia, mas conteúdo colado de fora costuma trazer `&nbsp;`.
_JS_BASE = """
const SELETORES = arguments[0];
function acharEditor() {
  for (const s of SELETORES) { const e = document.querySelector(s); if (e) return e; }
  return null;
}
function ehVazio(no) {
  return !!no && no.tagName === 'P' &&
         no.textContent.replace(/[\\s\\u200b\\u00a0]/g, '') === '';
}
function ehHeading(no) { return !!no && /^H[1-6]$/.test(no.tagName); }
// Percorre os blocos na ordem e devolve os índices dos parágrafos vazios cujo
// próximo bloco NÃO-vazio é um heading. Uma sequência de N vazios antes de um
// heading entra inteira; vazio antes de parágrafo comum é preservado.
function alvosDom(blocos) {
  const alvos = [];
  for (let i = 0; i < blocos.length; i++) {
    if (!ehVazio(blocos[i])) continue;
    let j = i + 1;
    while (j < blocos.length && ehVazio(blocos[j])) j++;
    if (j < blocos.length && ehHeading(blocos[j])) alvos.push(i);
  }
  return alvos;
}
"""

_JS_DIAGNOSTICO = _JS_BASE + """
const editor = acharEditor();
if (!editor) return {editor: false};
const blocos = Array.from(editor.children);
const container = editor.closest('.ql-container') || editor.parentElement;
let q = (container && container.__quill) || editor.__quill || null;
if (!q && window.Quill && window.Quill.find) {
  q = (container && window.Quill.find(container)) || window.Quill.find(editor) || null;
}
return {
  editor: true,
  quill: !!q,
  blocos: blocos.length,
  alvos: alvosDom(blocos).length,
  vazios_total: blocos.filter(ehVazio).length,
  texto: editor.textContent.replace(/[\\s\\u200b\\u00a0]/g, ''),
};
"""

# Caminho 1 — API do Quill. Remove de trás para frente: apagar a última linha
# primeiro mantém válidos os índices das linhas anteriores, então não é preciso
# recalcular nada a cada remoção.
_JS_LIMPAR_VIA_QUILL = _JS_BASE + """
const editor = acharEditor();
if (!editor) return {ok: false, motivo: 'sem_editor'};
const container = editor.closest('.ql-container') || editor.parentElement;
let q = (container && container.__quill) || editor.__quill || null;
if (!q && window.Quill && window.Quill.find) {
  q = (container && window.Quill.find(container)) || window.Quill.find(editor) || null;
}
if (!q) return {ok: false, motivo: 'sem_instancia'};

const linhas = q.getLines();
const doms = linhas.map(l => l.domNode);
const alvos = alvosDom(doms);

let removidos = 0;
for (let k = alvos.length - 1; k >= 0; k--) {
  const blot = q.getLines()[alvos[k]];
  if (!blot) continue;
  const idx = q.getIndex(blot);
  const tam = blot.length();          // parágrafo vazio = 1 (só o \\n)
  if (tam !== 1) continue;            // guarda: nunca apagar linha com conteúdo
  q.deleteText(idx, tam, 'user');     // 'user' → dispara onChange do React
  removidos++;
}
return {ok: true, removidos: removidos};
"""

# Caminho 2 — fallback por teclado. Posiciona o cursor no primeiro alvo e
# devolve quantos alvos ainda existem; o DELETE em si é enviado pelo Selenium
# (tecla real, não evento sintético, que o Quill ignoraria). A contagem volta
# junto para o laço perceber quando um DELETE não surtiu efeito, em vez de
# insistir até o teto de iterações.
_JS_POSICIONAR_CARET = _JS_BASE + """
const editor = acharEditor();
if (!editor) return {posicionou: false, alvos: 0};
const blocos = Array.from(editor.children);
const alvos = alvosDom(blocos);
if (!alvos.length) return {posicionou: false, alvos: 0};
const alvo = blocos[alvos[0]];
alvo.scrollIntoView({block: 'center'});
editor.focus();
const faixa = document.createRange();
faixa.setStart(alvo, 0);
faixa.collapse(true);
const selecao = window.getSelection();
selecao.removeAllRanges();
selecao.addRange(faixa);
return {posicionou: true, alvos: alvos.length};
"""


def _diagnosticar(driver, seletores):
    return driver.execute_script(_JS_DIAGNOSTICO, seletores)


def _esperar_editor(driver, seletores, log):
    """Espera o componente de SEO montar e o conteúdo chegar; devolve o
    diagnóstico (ou None se o navegador nem respondeu).

    Duas etapas, porque são dois atrasos diferentes:

      1. o React ainda está montando o formulário — `#seo .ql-editor` sequer
         existe no DOM (era aqui que o módulo desistia, milissegundos depois de
         a coleção abrir);
      2. o editor já existe mas está vazio, esperando o texto vir do backend.
         Começar a apagar linha nesse meio-tempo faria a conferência de
         integridade comparar textos de momentos diferentes e abortar o fluxo à
         toa — por isso o texto também precisa parar de mudar entre duas
         leituras antes de a limpeza começar.

    Editor ausente depois de TIMEOUT_EDITOR ou vazio depois de TIMEOUT_CONTEUDO
    é resposta legítima (tela sem SEO / coleção sem texto): devolve o último
    diagnóstico e quem chamou apenas loga e segue.
    """
    def ler():
        try:
            return _diagnosticar(driver, seletores)
        except Exception as e:
            log.warning(f"  ⚠ Não foi possível inspecionar o editor de SEO: {e}")
            return None

    fim = time.time() + TIMEOUT_EDITOR
    diagnostico = ler()
    while diagnostico is not None and not diagnostico.get('editor'):
        if time.time() >= fim:
            return diagnostico
        time.sleep(INTERVALO_POLL)
        diagnostico = ler()

    if diagnostico is None or not diagnostico.get('editor'):
        return diagnostico

    fim = time.time() + TIMEOUT_CONTEUDO
    while diagnostico is not None and not diagnostico.get('texto'):
        if time.time() >= fim:
            return diagnostico
        time.sleep(INTERVALO_POLL)
        diagnostico = ler()

    # Conteúdo estável: duas leituras seguidas com o mesmo texto. Evita pegar o
    # editor no meio do preenchimento e acusar "o texto mudou" mais adiante.
    fim = time.time() + TIMEOUT_CONTEUDO
    while diagnostico is not None and diagnostico.get('texto'):
        time.sleep(INTERVALO_POLL)
        confirmacao = ler()
        if confirmacao is None or confirmacao.get('texto') == diagnostico.get('texto'):
            return confirmacao if confirmacao is not None else diagnostico
        diagnostico = confirmacao
        if time.time() >= fim:
            break

    return diagnostico


def limpar_paragrafos_vazios_seo(driver, seletores=None, log=None):
    """Remove os parágrafos vazios que o editor injeta antes dos headings.

    Deve ser chamada com a tela de edição já aberta e ANTES de qualquer
    alteração nos produtos, para que o "Salvar" no fim do fluxo persista o
    conteúdo já corrigido.

    Devolve um dict:
        {'editor_encontrado', 'removidos', 'restantes', 'metodo',
         'texto_preservado'}
    Nunca levanta exceção: qualquer falha é logada e devolvida no dict, para
    que a atualização de produtos não trave por causa do SEO.
    """
    log = log or logger
    seletores = seletores or SELETORES_EDITOR
    resultado = {
        'editor_encontrado': False,
        'removidos': 0,
        'restantes': 0,
        'metodo': None,
        'texto_preservado': True,
    }

    antes = _esperar_editor(driver, seletores, log)

    if not antes or not antes.get('editor'):
        log.info(
            f"Editor de SEO não encontrado nesta tela após {TIMEOUT_EDITOR}s "
            "— nada a limpar"
        )
        return resultado

    resultado['editor_encontrado'] = True
    texto_original = antes.get('texto', '')

    if not texto_original:
        log.info(
            f"Editor de SEO continuou vazio após {TIMEOUT_CONTEUDO}s — nada a limpar"
        )
        return resultado

    alvos = antes.get('alvos', 0)
    log.info(
        f"Editor de SEO: {antes.get('blocos', 0)} blocos, "
        f"{antes.get('vazios_total', 0)} parágrafo(s) vazio(s), "
        f"{alvos} antes de heading"
    )

    if not alvos:
        log.info("✓ Estrutura do SEO já está correta")
        return resultado

    # --- Caminho 1: API do Quill ---
    if antes.get('quill'):
        try:
            via_quill = driver.execute_script(_JS_LIMPAR_VIA_QUILL, seletores)
            if via_quill and via_quill.get('ok'):
                resultado['metodo'] = 'quill'
                resultado['removidos'] = via_quill.get('removidos', 0)
            else:
                motivo = (via_quill or {}).get('motivo', 'desconhecido')
                log.info(f"  Instância do Quill indisponível ({motivo}) — usando teclado")
        except Exception as e:
            log.warning(f"  ⚠ Falha na limpeza via API do Quill: {e} — usando teclado")

    # --- Caminho 2: fallback por teclado (roda se o Quill não resolveu tudo) ---
    if resultado['metodo'] == 'quill':
        try:
            sobraram = (_diagnosticar(driver, seletores) or {}).get('alvos', 0)
        except Exception as e:
            log.warning(f"  ⚠ Não foi possível reconferir o editor após o Quill: {e}")
            sobraram = 0
    else:
        sobraram = True

    if sobraram:
        removidos_teclado = _limpar_via_teclado(driver, seletores, log)
        if removidos_teclado:
            resultado['metodo'] = 'teclado' if not resultado['metodo'] else 'quill+teclado'
            resultado['removidos'] += removidos_teclado

    # --- Conferência final ---
    # O React ainda pode estar re-renderizando o editor logo depois da última
    # remoção; ler antes disso mostraria texto pela metade.
    time.sleep(INTERVALO_POLL)
    try:
        depois = _diagnosticar(driver, seletores)
        if depois and depois.get('texto', '') != texto_original:
            # Segunda chance antes de acusar perda de conteúdo: um falso
            # positivo aqui aborta o fluxo inteiro por nada.
            time.sleep(INTERVALO_POLL * 4)
            depois = _diagnosticar(driver, seletores)
    except Exception as e:
        log.warning(f"  ⚠ Não foi possível conferir o editor após a limpeza: {e}")
        return resultado

    if not depois:
        log.warning("  ⚠ Não foi possível conferir o editor após a limpeza")
        return resultado

    resultado['restantes'] = depois.get('alvos', 0)

    # O texto visível (sem espaços) tem que ser exatamente o mesmo: só linhas
    # vazias foram removidas. Se mudou, algo apagou conteúdo de verdade.
    if depois.get('texto', '') != texto_original:
        resultado['texto_preservado'] = False
        log.error(
            "  ✗ O TEXTO DO SEO MUDOU durante a limpeza — confira a página "
            "manualmente ANTES de salvar"
        )
        return resultado

    if resultado['restantes']:
        log.warning(
            f"  ⚠ Ainda restam {resultado['restantes']} parágrafo(s) vazio(s) "
            f"antes de heading (removidos: {resultado['removidos']})"
        )
    else:
        log.info(
            f"✓ SEO normalizado: {resultado['removidos']} parágrafo(s) vazio(s) "
            f"removido(s) via {resultado['metodo']}"
        )

    return resultado


def _limpar_via_teclado(driver, seletores, log):
    """Posiciona o cursor no parágrafo vazio e manda DELETE, um por vez.

    O DOM é relido a cada iteração de propósito: o Quill reconstrói os nós
    depois de cada edição, então guardar referência de elemento entre iterações
    daria StaleElementReference.
    """
    removidos = 0
    alvos_anteriores = None
    sem_progresso = 0

    for _ in range(MAX_ITERACOES_TECLADO):
        try:
            estado = driver.execute_script(_JS_POSICIONAR_CARET, seletores)
        except Exception as e:
            log.warning(f"  ⚠ Falha ao posicionar o cursor no editor: {e}")
            break

        if not estado or not estado.get('posicionou'):
            break  # não há mais alvo

        alvos = estado.get('alvos', 0)
        if alvos_anteriores is not None:
            if alvos < alvos_anteriores:
                removidos += alvos_anteriores - alvos
                sem_progresso = 0
            else:
                # O DELETE anterior não tirou nada: ou o Quill ainda não
                # sincronizou, ou a tecla não chegou ao editor. Espera um pouco
                # mais antes da próxima e desiste depois de algumas tentativas,
                # em vez de girar até MAX_ITERACOES_TECLADO.
                sem_progresso += 1
                if sem_progresso >= MAX_TENTATIVAS_SEM_PROGRESSO:
                    log.warning(
                        "  ⚠ O DELETE não surtiu efeito no editor de SEO — "
                        f"parando com {alvos} parágrafo(s) ainda por remover"
                    )
                    break
                time.sleep(PAUSA_ENTRE_TECLAS * 5)
        alvos_anteriores = alvos

        try:
            # ActionChains nova a cada volta: reaproveitar a mesma instância
            # pode acumular ações já executadas dependendo da versão do
            # Selenium, e um DELETE a mais apagaria conteúdo de verdade.
            ActionChains(driver).send_keys(Keys.DELETE).perform()
        except Exception as e:
            log.warning(f"  ⚠ Falha ao enviar DELETE ao editor: {e}")
            break

        time.sleep(PAUSA_ENTRE_TECLAS)
    else:
        log.warning(
            f"  ⚠ Limite de {MAX_ITERACOES_TECLADO} remoções atingido — "
            "possível loop; confira a página"
        )

    # O laço sai logo após um DELETE, sem uma nova contagem: confere quanto
    # sobrou para creditar a última remoção (e só ela).
    if alvos_anteriores:
        try:
            restantes = driver.execute_script(_JS_POSICIONAR_CARET, seletores)
            removidos += max(0, alvos_anteriores - (restantes or {}).get('alvos', 0))
        except Exception:
            pass

    return removidos
