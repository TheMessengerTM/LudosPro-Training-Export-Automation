"""
Extrai o conteúdo dos treinamentos do LudosPro e gera um documento Word.

Para cada curso:  título + descrição (aba Curso)
  -> cada módulo -> cada atividade -> cada mecânica Slideshow (abre "Editar" só para ler)
     -> título, introdução e título + descrição de cada slide.
  As demais mecânicas (Quiz, This or That etc.) são ignoradas.

NUNCA salva nada no LudosPro: toda edição aberta é fechada com "Cancelar".

Uso:
    python Src/ludos_treinamentos.py              # todos os cursos
    python Src/ludos_treinamentos.py 138 137      # só os cursos com esses IDs
    python Src/ludos_treinamentos.py --refazer    # ignora o cache e extrai tudo de novo
    python Src/ludos_treinamentos.py --oculto     # Chrome sem janela (usa menos memória; exige login já salvo)

Configuração (.env na raiz do projeto ou variáveis de ambiente):
    LUDOS_URL   endereço do painel administrativo (ex.: https://sua-empresa.ludos.pro/adm3/)
    LUDOS_USER  e LUDOS_PASS: opcionais. Se estiverem definidos, o login é automático;
                senão, faça o login na janela e tecle Enter no terminal.
A sessão fica guardada na pasta .chrome_ludos.

Saída:
    saida/cache/<id>.json            dados de cada curso (permite retomar)
    saida/Treinamentos_LudosPro.docx documento final
"""
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt, RGBColor
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

PASTA = Path(__file__).resolve().parent.parent


def carregar_env(caminho=PASTA / ".env"):
    """Lê pares CHAVE=VALOR do arquivo .env para as variáveis de ambiente."""
    if not caminho.exists():
        return
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip())


carregar_env()
BASE = os.environ.get("LUDOS_URL", "").rstrip("/") + "/"
PERFIL = PASTA / ".chrome_ludos"
SAIDA = PASTA / "saida"
CACHE = SAIDA / "cache"
ARQ_WORD = SAIDA / "Treinamentos_LudosPro.docx"


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# --------------------------------------------------------------------------- navegação
def ir(page, rota):
    page.evaluate("r => { location.hash = r }", rota)
    page.wait_for_timeout(800)


_pendentes = set()


def monitorar_rede(page):
    """Conta as requisições XHR/fetch em andamento, para saber quando a tela terminou de carregar."""
    def inicio(r):
        if r.resource_type in ("xhr", "fetch"):
            _pendentes.add(r)

    def fim(r):
        _pendentes.discard(r)

    page.on("request", inicio)
    page.on("requestfinished", fim)
    page.on("requestfailed", fim)


def esperar_rede(page, quieto=0.6, timeout=30.0):
    """Espera não haver requisições pendentes por `quieto` segundos seguidos."""
    page.wait_for_timeout(300)  # dá tempo da requisição disparar
    fim = time.time() + timeout
    desde = None
    while time.time() < fim:
        if _pendentes:
            desde = None
        elif desde is None:
            desde = time.time()
        elif time.time() - desde >= quieto:
            break
        page.wait_for_timeout(100)
    else:
        _pendentes.clear()  # estourou o tempo: descarta requisições presas (ex.: polling)


def esperar_estavel(page, seletor, timeout=10.0, intervalo=0.5):
    """Espera a rede acalmar e o texto de `seletor` parar de mudar."""
    esperar_rede(page)
    fim = time.time() + timeout
    anterior = None
    while time.time() < fim:
        page.wait_for_timeout(int(intervalo * 1000))
        atual = page.evaluate(
            "s => [...document.querySelectorAll(s)].map(e => e.innerText).join('|')", seletor
        )
        if atual == anterior:
            return
        anterior = atual


def garantir_login(page):
    page.goto(BASE + "#/courses/personal")
    page.wait_for_timeout(3000)
    if "/login" not in page.url:
        return
    usuario, senha = os.environ.get("LUDOS_USER"), os.environ.get("LUDOS_PASS")
    if usuario and senha:
        log("Fazendo login...")
        page.fill("#app-input-0", usuario)
        page.fill("#app-input-password-0", senha)
        page.click("button[type=submit]")
    else:
        input(">> Faça o login no LudosPro na janela do Chrome e tecle Enter aqui... ")
    page.wait_for_url(lambda u: "/login" not in u, timeout=120_000)
    page.wait_for_timeout(2000)


def listar_cursos(page):
    ir(page, "#/courses/personal")
    page.wait_for_selector("app-courses-card", timeout=30_000)
    # mostra 96 por página
    try:
        page.locator("app-pager app-select button").first.click()
        page.locator(".cdk-overlay-container").get_by_text("96", exact=True).click()
        esperar_estavel(page, "app-courses-card")
    except Exception:
        pass
    cursos = {}
    while True:
        for a in page.locator("a.c-card-info__title").all():
            m = re.search(r"courses/(\d+)/", a.get_attribute("href") or "")
            if m:
                cursos[int(m.group(1))] = a.inner_text().strip()
        prox = page.locator("app-pager .c-navigate--next")
        if prox.count() == 0 or "disabled" in (prox.get_attribute("class") or ""):
            break
        prox.click()
        esperar_estavel(page, "app-courses-card")
    return cursos


JS_CAMPO = """sel => {
  const f = document.querySelector(sel);
  if (!f) return '';
  const q = f.querySelector('.ql-editor');
  if (q) return q.innerText;
  const i = f.matches('input,textarea') ? f : f.querySelector('input,textarea');
  return i ? i.value : '';
}"""


def ler_aba_curso(page, cid, titulo_esperado=""):
    ir(page, f"#/courses/{cid}/update")
    esperar_rede(page)
    # o formulário pode mostrar por um instante o curso anterior: espera aparecer o título
    # que a lista de cursos indica para este ID
    js = """esp => { const i = document.querySelector('[formcontrolname=txCourse] input, input[formcontrolname=txCourse]');
             const n = s => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
             return i && n(i.value).length > 0 && (!esp || n(i.value) === n(esp)) }"""
    try:
        page.wait_for_function(js, arg=titulo_esperado, timeout=30_000)
    except PWTimeout:
        log(f"   ! título da aba Curso não bateu com a lista (esperado: {titulo_esperado})")
        page.wait_for_function(js, arg="", timeout=10_000)
    esperar_rede(page)
    page.wait_for_timeout(500)
    return {
        "titulo": page.evaluate(JS_CAMPO, "[formcontrolname=txCourse]").strip(),
        "descricao": page.evaluate(JS_CAMPO, "[formcontrolname=txDescription]").strip(),
    }


JS_MODAL = """() => {
  const m = document.querySelector('.cdk-overlay-pane app-modal');
  if (!m) return null;
  const v = s => { const e = m.querySelector(s); return e ? (e.value ?? e.innerText ?? '') : ''; };
  return {
    tipo_modal: m.querySelector('h1')?.innerText || '',
    titulo: v('app-input[formcontrolname=tx_titulo] input'),
    introducao: v('app-textarea[formcontrolname=tx_introducao] textarea'),
    paineis: [...m.querySelectorAll('.--panel-mecanica')].map(p => ({
      cabecalho: p.querySelector('h4')?.innerText || '',
      titulo: p.querySelector('[formcontrolname=tx_titulo] .ql-editor')?.innerText || '',
      descricao: p.querySelector('[formcontrolname=tx_slide] .ql-editor')?.innerText || '',
    })),
  };
}"""


def fechar_modal(page):
    modal = page.locator(".cdk-overlay-pane app-modal")
    if modal.count() == 0:
        return
    cancelar = modal.get_by_role("button", name="Cancelar", exact=True)
    if cancelar.count():
        cancelar.first.click()
    else:
        modal.locator(".--header-mecanica__close").first.click()
    page.wait_for_function(
        "() => !document.querySelector('.cdk-overlay-pane app-modal')", timeout=10_000
    )
    page.wait_for_timeout(300)


def abrir_editar(page, card):
    """Abre o menu ⋮ do card e clica SOMENTE em 'Editar'."""
    card.locator(".c-card__menu-button").first.click()
    item = page.locator(".cdk-overlay-pane").get_by_text("Editar", exact=True)
    item.first.wait_for(state="visible", timeout=5_000)
    item.first.click()


def ler_mecanica(page, card):
    abrir_editar(page, card)
    page.wait_for_selector(".cdk-overlay-pane app-modal", timeout=15_000)
    # espera o título carregar no formulário
    try:
        page.wait_for_function(
            "() => { const i = document.querySelector('.cdk-overlay-pane app-modal app-input[formcontrolname=tx_titulo] input');"
            " return i && i.value.trim().length > 0 }",
            timeout=10_000,
        )
    except PWTimeout:
        pass
    page.wait_for_timeout(700)
    dados = page.evaluate(JS_MODAL)
    fechar_modal(page)
    return dados


def extrair_curso(page, cid, titulo_esperado=""):
    curso = ler_aba_curso(page, cid, titulo_esperado)
    curso["cabecalho_ok"] = True
    curso["id"] = cid
    curso["modulos"] = []

    ir(page, f"#/courses/{cid}/content")
    try:
        page.wait_for_selector("app-panel-module", timeout=20_000)
    except PWTimeout:
        log("   (curso sem módulos)")
        return curso
    esperar_estavel(page, "app-panel-module")

    n_mod = page.locator("app-panel-module").count()
    for i in range(n_mod):
        mod_card = page.locator("app-panel-module").nth(i)
        nome_mod = mod_card.locator(".c-card__text").first.get_attribute("title") or mod_card.inner_text()
        log(f"   Módulo {i+1}/{n_mod}: {nome_mod.strip()}")
        mod_card.locator(".c-card__text").first.click()
        page.wait_for_timeout(800)
        esperar_estavel(page, "app-panel-activity")
        modulo = {"nome": nome_mod.strip(), "atividades": []}

        n_at = page.locator("app-panel-activity").count()
        for j in range(n_at):
            at_card = page.locator("app-panel-activity").nth(j)
            nome_at = at_card.locator(".c-card__text").first.get_attribute("title") or at_card.inner_text()
            at_card.locator(".c-card__text").first.click()
            page.wait_for_timeout(800)
            esperar_estavel(page, "app-panel-game")
            atividade = {"titulo": nome_at.strip(), "mecanicas": []}

            n_g = page.locator("app-panel-game").count()
            for k in range(n_g):
                g_card = page.locator("app-panel-game").nth(k)
                tipo = (g_card.locator(".c-card__game-type").first.get_attribute("title") or "").strip()
                if tipo.lower() != "slideshow":
                    continue  # só interessam os Slideshows
                nome_g = (g_card.locator(".c-card__game-name").first.get_attribute("title") or "").strip()
                mec = {"tipo": tipo, "titulo": nome_g, "introducao": "", "slides": []}
                try:
                    d = ler_mecanica(page, g_card)
                    if d:
                        mec["titulo"] = (d["titulo"] or nome_g).strip()
                        mec["introducao"] = (d["introducao"] or "").strip()
                        if tipo.lower() == "slideshow":
                            mec["slides"] = [
                                {"titulo": p["titulo"].strip() or re.sub(r"^\d+\.\s*", "", p["cabecalho"]).strip(),
                                 "descricao": p["descricao"].strip()}
                                for p in d["paineis"]
                            ]
                except Exception as e:  # registra e segue para a próxima
                    mec["erro"] = str(e).splitlines()[0]
                    log(f"      ! erro na mecânica '{nome_g}': {mec['erro']}")
                    try:
                        page.keyboard.press("Escape")
                        fechar_modal(page)
                    except Exception:
                        pass
                atividade["mecanicas"].append(mec)
            modulo["atividades"].append(atividade)
        curso["modulos"].append(modulo)
    return curso


# --------------------------------------------------------------------------- Word
_INVALIDOS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def limpo(t):
    t = _INVALIDOS.sub("", t or "")
    return re.sub(r"\n{3,}", "\n\n", t).strip()


def gerar_word(cursos, destino):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(11)

    doc.add_heading("Treinamentos LudosPro", 0)
    doc.add_paragraph(f"Extraído em {datetime.now():%d/%m/%Y %H:%M} — {len(cursos)} treinamento(s).")

    cinza = RGBColor(0x59, 0x59, 0x59)
    for n, c in enumerate(cursos):
        if n:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        doc.add_heading(f"Treinamento: {limpo(c['titulo'])}", 1)
        p = doc.add_paragraph()
        p.add_run("Descrição: ").bold = True
        p.add_run(limpo(c.get("descricao")) or "—")
        if c.get("erro"):
            doc.add_paragraph(f"[ERRO NA EXTRAÇÃO: {c['erro']}]")

        for mi, m in enumerate(c.get("modulos", []), 1):
            doc.add_heading(f"Módulo {mi}: {limpo(m['nome'])}", 2)
            for ai, a in enumerate(m["atividades"], 1):
                doc.add_heading(f"Atividade {ai}: {limpo(a['titulo'])}", 3)
                slideshows = [g for g in a["mecanicas"] if g["tipo"].lower() == "slideshow"]
                if not slideshows:
                    doc.add_paragraph("(sem slideshow)")
                for gi, g in enumerate(slideshows, 1):
                    doc.add_heading(f"Slideshow {gi}: {limpo(g['titulo'])}", 4)
                    if g.get("introducao"):
                        p = doc.add_paragraph()
                        p.add_run("Introdução: ").bold = True
                        p.add_run(limpo(g["introducao"]))
                    if g.get("erro"):
                        doc.add_paragraph(f"[não foi possível abrir a edição: {g['erro']}]")
                    for si, s in enumerate(g.get("slides", []), 1):
                        p = doc.add_paragraph()
                        p.paragraph_format.left_indent = Pt(18)
                        p.paragraph_format.space_after = Pt(2)
                        p.add_run(f"Slide {si} – {limpo(s['titulo'])}").bold = True
                        if s.get("descricao"):
                            p = doc.add_paragraph()
                            p.paragraph_format.left_indent = Pt(18)
                            r = p.add_run(limpo(s["descricao"]))
                            r.font.color.rgb = cinza
    destino.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(destino)
    except PermissionError:  # arquivo aberto no Word: salva uma cópia ao lado
        alt = destino.with_name(destino.stem + "_novo" + destino.suffix)
        try:
            doc.save(alt)
            log(f"   (Word aberto — salvo em {alt.name}; feche o Word para atualizar o original)")
        except PermissionError:
            log("   (Word aberto — não foi possível salvar agora; tentarei no próximo curso)")


# --------------------------------------------------------------------------- main
def main():
    if BASE == "/":
        raise SystemExit("Defina LUDOS_URL no arquivo .env (veja o .env.example)")
    args = sys.argv[1:]
    refazer = "--refazer" in args
    cdp = "--cdp" in args  # uso interno: conecta num Chrome já aberto na porta 9222
    ids_pedidos = [int(a) for a in args if a.isdigit()]
    CACHE.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        if cdp:
            browser = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            ctx = browser.contexts[0]
            page = ctx.pages[-1]
        else:
            ctx = p.chromium.launch_persistent_context(
                str(PERFIL), channel="chrome", headless="--oculto" in args,
                viewport={"width": 1400, "height": 900}
            )
            page = ctx.pages[0] if ctx.pages else ctx.new_page()

        monitorar_rede(page)
        garantir_login(page)
        todos = listar_cursos(page)
        log(f"{len(todos)} cursos encontrados.")
        ids = ids_pedidos or list(todos)

        cursos = []
        for idx, cid in enumerate(ids, 1):
            arq = CACHE / f"{cid}.json"
            if arq.exists() and not refazer:
                c = json.loads(arq.read_text(encoding="utf-8"))
                if not c.get("cabecalho_ok"):
                    # cache antigo: o título/descrição podem ter vindo do curso anterior — relê só a aba Curso
                    log(f"[{idx}/{len(ids)}] {cid} em cache — relendo título e descrição")
                    try:
                        c.update(ler_aba_curso(page, cid, todos.get(cid, "")))
                        c["cabecalho_ok"] = True
                        arq.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")
                    except Exception as e:
                        log(f"   ! não consegui reler a aba Curso: {str(e).splitlines()[0]}")
                else:
                    log(f"[{idx}/{len(ids)}] {cid} já extraído (cache) — pulando")
                cursos.append(c)
                continue
            log(f"[{idx}/{len(ids)}] Curso {cid}: {todos.get(cid, '')}")
            try:
                c = extrair_curso(page, cid, todos.get(cid, ""))
                arq.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")
            except Exception as e:
                log(f"   ! erro no curso {cid}: {e}")
                c = {"id": cid, "titulo": todos.get(cid, str(cid)), "descricao": "", "modulos": [],
                     "erro": str(e).splitlines()[0]}
            cursos.append(c)
            gerar_word(cursos, ARQ_WORD)  # salva parcial a cada curso

        # o Word final sempre traz todos os cursos (na ordem da lista), mesmo rodando só alguns IDs
        extraidos = {c["id"]: c for c in cursos}
        final = []
        for cid in todos:
            if cid in extraidos:
                final.append(extraidos[cid])
            elif (CACHE / f"{cid}.json").exists():
                final.append(json.loads((CACHE / f"{cid}.json").read_text(encoding="utf-8")))
        gerar_word(final, ARQ_WORD)
        log(f"Pronto! Word salvo em: {ARQ_WORD} ({len(final)} cursos)")
        if not cdp:
            ctx.close()


if __name__ == "__main__":
    main()
