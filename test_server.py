"""Auto-teste: python test_server.py"""
import base64, json, os, subprocess, sys, tempfile, time, urllib.error, urllib.request
here = os.path.dirname(os.path.abspath(__file__))
d = os.path.join(tempfile.mkdtemp(), "F0000TEST"); os.makedirs(d)
open(os.path.join(d, "ficha.pdf"), "wb").write(b"%PDF-1.4 fake")
# como o VBA grava: ANSI, com acento
open(os.path.join(d, "slots.json"), "w", encoding="cp1252").write(
    '[["offset","Impressão Offset",0.0,0.9,0.2,0.02],["pcp","PCP",0.3,0.95,0.15,0.02]]')
local = "http://127.0.0.1:5099"


def sobe(pasta):
    for f in ("url.txt", "stop.flag"):
        if os.path.exists(os.path.join(pasta, f)): os.remove(os.path.join(pasta, f))
    proc = subprocess.Popen([sys.executable, os.path.join(here, "server.py"), pasta, "5099"])
    for _ in range(50):
        if os.path.exists(os.path.join(pasta, "url.txt")): break
        time.sleep(.1)
    return proc, open(os.path.join(pasta, "url.txt")).read().split("/?t=")[1]


def req(path, data=None, headers={}):
    try:
        r = urllib.request.urlopen(urllib.request.Request(local + path, data=data, headers=headers))
        return r.status, r.read()
    except urllib.error.HTTPError as e: return e.code, b""


st = lambda t: json.loads(req("/status?t=" + t)[1])
p, t = sobe(d)
try:
    png = b"\x89PNG\r\n\x1a\n" + b"x" * 20
    b64 = ("data:image/png;base64," + base64.b64encode(png).decode()).encode()
    assert req("/")[0] == 403, "sem token deve negar"
    code, html = req("/?t=" + t)
    assert code == 200 and "Impress\\u00e3o Offset" in html.decode(), "titulo com acento chega na pagina"
    assert len(t) == 6 and not set(t) & set("l1o0"), "token digitavel"
    assert req("/qr")[0] == 403 and req("/qr?t=" + t)[0] == 200
    assert req("/favicon.ico")[0] == 204
    assert req("/ficha.png?t=" + t)[0] == 404, "sem imagem -> 404, pagina mostra aviso"
    assert req("/save/offset?t=wrong", b64)[0] == 403
    assert req("/save/eng?t=" + t, b64)[0] == 403, "setor fora do slots.json"
    assert req("/save/offset?t=" + t, b"data:image/png;base64,AAAA")[0] == 400, "nao-PNG"
    assert req("/sig/offset.png?t=" + t)[0] == 404
    assert req("/save/offset?t=" + t, b64)[0] == 200
    assert req("/sig/offset.png?t=" + t) == (200, png)
    assert req("/sig/..%2Fslots.json?t=" + t)[0] == 404, "sem path traversal"
    assert st(t) == {"done": ["offset"], "fase": "coletando", "erro": ""}

    # retomada: servidor cai e sobe de novo -> mesmo link, assinatura anterior continua
    open(os.path.join(d, "stop.flag"), "w").close(); p.wait(timeout=5)
    p, t2 = sobe(d)
    assert t2 == t, "link do celular deve continuar valendo depois de reiniciar"
    assert st(t)["done"] == ["offset"]
    assert req("/save/pcp?t=" + t, b64, {"X-Modo": "digitado"})[0] == 200

    # segunda coleta na mesma porta deve falhar (sem url.txt), nao dividir a porta
    d2 = os.path.join(tempfile.mkdtemp(), "F0000DUP"); os.makedirs(d2)
    p2 = subprocess.run([sys.executable, os.path.join(here, "server.py"), d2, "5099"], capture_output=True, timeout=10)
    assert p2.returncode != 0 and not os.path.exists(os.path.join(d2, "url.txt")), "porta deve ser exclusiva"

    # finalizar pelo celular: flag para o Excel; erro do Excel aparece; depois concluido com o PDF
    assert req("/finalizar?t=wrong", b"")[0] == 403
    assert req("/finalizar?t=" + t, b"")[0] == 200
    assert os.path.exists(os.path.join(d, "finalizar.flag")) and st(t)["fase"] == "finalizando"
    os.remove(os.path.join(d, "finalizar.flag"))
    open(os.path.join(d, "erro.txt"), "w", encoding="cp1252").write("Planilha fechada")
    assert st(t)["fase"] == "erro" and st(t)["erro"] == "Planilha fechada"
    assert req("/finalizar?t=" + t, b"")[0] == 200 and st(t)["fase"] == "finalizando", "pode tentar de novo"
    assert not os.path.exists(os.path.join(d, "erro.txt"))
    open(os.path.join(d, "pdf.txt"), "w").write(r"X:\Assinadas\F0000TEST_assinada_20260101_120000.pdf")
    open(os.path.join(d, "stop.flag"), "w").close()
    time.sleep(1.6)
    os.rename(d, d + "_arquivada")  # o macro arquiva a pasta enquanto o celular ainda consulta
    s = st(t)
    assert s["fase"] == "concluido" and s["pdf"] == "F0000TEST_assinada_20260101_120000.pdf", s
    assert sorted(s["done"]) == ["offset", "pcp"]
    p.wait(timeout=12)
    log = open(os.path.join(d + "_arquivada", "log.csv"), encoding="utf-8").read()
    assert ",offset,desenho" in log and ",pcp,digitado" in log and ",finalizar,celular" in log
    print("OK: todos os testes passaram; token =", t)
finally:
    if p.poll() is None: p.kill()
