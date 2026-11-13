"""Auto-teste: python test_server.py"""
import base64, os, subprocess, sys, tempfile, time, urllib.error, urllib.request
here = os.path.dirname(os.path.abspath(__file__))
d = os.path.join(tempfile.mkdtemp(), "F0000TEST"); os.makedirs(d)
open(os.path.join(d, "ficha.pdf"), "wb").write(b"%PDF-1.4 fake")
# como o VBA grava: ANSI, com acento
open(os.path.join(d, "slots.json"), "w", encoding="cp1252").write(
    '[["offset","Impressão Offset",0.0,0.9,0.2,0.02],["pcp","PCP",0.3,0.95,0.15,0.02]]')
p = subprocess.Popen([sys.executable, os.path.join(here, "server.py"), d, "5099"])
try:
    for _ in range(50):
        if os.path.exists(os.path.join(d, "url.txt")): break
        time.sleep(.1)
    url = open(os.path.join(d, "url.txt")).read(); t = url.split("/?t=")[1]
    local = "http://127.0.0.1:5099"
    def req(path, data=None, headers={}):
        try:
            r = urllib.request.urlopen(urllib.request.Request(local + path, data=data, headers=headers))
            return r.status, r.read()
        except urllib.error.HTTPError as e: return e.code, b""
    png = b"\x89PNG\r\n\x1a\n" + b"x" * 20
    b64 = ("data:image/png;base64," + base64.b64encode(png).decode()).encode()
    assert req("/")[0] == 403, "sem token deve negar"
    code, html = req("/?t=" + t)
    assert code == 200 and "Impress\\u00e3o Offset" in html.decode(), "titulo com acento chega na pagina"
    assert len(t) == 6 and not set(t) & set("l1o0"), "token digitavel"
    assert req("/qr")[0] == 403 and req("/qr?t=" + t)[0] == 200
    assert req("/ficha.png?t=" + t)[0] == 404, "sem imagem -> 404, pagina mostra aviso"
    assert req("/save/offset?t=wrong", b64)[0] == 403
    assert req("/save/eng?t=" + t, b64)[0] == 403, "setor fora do slots.json"
    assert req("/save/offset?t=" + t, b"data:image/png;base64,AAAA")[0] == 400, "nao-PNG"
    assert req("/sig/offset.png?t=" + t)[0] == 404
    assert req("/save/offset?t=" + t, b64)[0] == 200
    assert req("/save/pcp?t=" + t, b64, {"X-Modo": "digitado"})[0] == 200
    assert req("/sig/offset.png?t=" + t) == (200, png)
    assert req("/sig/..%2Fslots.json?t=" + t)[0] == 404, "sem path traversal"
    assert sorted(eval(req("/status?t=" + t)[1])) == ["offset", "pcp"]
    log = open(os.path.join(d, "log.csv"), encoding="utf-8").read()
    assert ",offset,desenho" in log and ",pcp,digitado" in log
    # segunda coleta na mesma porta deve falhar (sem url.txt), nao dividir a porta
    d2 = os.path.join(tempfile.mkdtemp(), "F0000DUP"); os.makedirs(d2)
    p2 = subprocess.run([sys.executable, os.path.join(here, "server.py"), d2, "5099"], capture_output=True, timeout=10)
    assert p2.returncode != 0 and not os.path.exists(os.path.join(d2, "url.txt")), "porta deve ser exclusiva"
    open(os.path.join(d, "stop.flag"), "w").close()
    p.wait(timeout=5)
    print("OK: todos os testes passaram; url =", url)
finally:
    if p.poll() is None: p.kill()
