"""Mini servidor de coleta de assinaturas da ficha da faca. So biblioteca padrao.

Uso: python server.py <pasta_da_coleta> [porta]
A pasta deve conter ficha.pdf, ficha.png e slots.json (gerados pelo macro).
Cria assinaturas/<setor>.png, log.csv e url.txt. Encerra quando aparecer stop.flag na pasta.
"""
import base64, csv, datetime, json, os, secrets, socket, sys, threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PNG = b"\x89PNG\r\n\x1a\n"
MAX_BYTES = 3_000_000
ALFA = "abcdefghjkmnpqrstuvwxyz23456789"  # sem l/1/o/0: o token pode ser digitado
MODOS = ("desenho", "digitado")
SLOTS_PADRAO = [
    ["offset", "Impressão Offset", 0.0, 0.8867, 0.2013, 0.0322], ["corte", "Corte e Vinco", 0.2827, 0.8867, 0.1628, 0.0322],
    ["colagem", "Colagem/Destaque", 0.527, 0.8867, 0.1628, 0.0322], ["eng", "Eng. Processos", 0.7712, 0.8867, 0.1474, 0.0322],
    ["coord", "Coordenador de Produção", 0.0, 0.9302, 0.2013, 0.0358], ["pcp", "PCP", 0.2827, 0.9302, 0.1628, 0.0358],
    ["excelencia", "Excelência Operacional", 0.527, 0.9302, 0.1628, 0.0358],
]

PAGE = """<!doctype html><html lang=pt-br><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name=theme-color content="#13293d">
<title>Assinaturas __FACA__</title>
<style>
/* Ficha da faca no chao de fabrica: leitura rapida, alvos grandes, papel claro.
   Destaque unico: o codigo da faca no amarelo da ficha impressa. */
:root{--paper:#edf0f2;--card:#fff;--ink:#13293d;--muted:#566676;--line:#d0d8de;--accent:#1d6fb8;
 --pen:#0b2a6b;--ok:#1e7a46;--okbg:#e2f2e7;--okline:#b9dfc6;--code:#fff59a;--warn:#8a1c17;--warnbg:#fde2e1;
 --sl:env(safe-area-inset-left,0px);--sr:env(safe-area-inset-right,0px);--st:env(safe-area-inset-top,0px);--sb:env(safe-area-inset-bottom,0px);
 color-scheme:light}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);font:16px/1.4 "Segoe UI",Roboto,system-ui,-apple-system,sans-serif;-webkit-tap-highlight-color:transparent}
button{font:inherit;color:inherit;cursor:pointer}
:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
[hidden]{display:none!important}

.top{position:sticky;top:0;z-index:3;background:var(--ink);color:#fff;
 padding:calc(10px + var(--st)) max(16px,var(--sr)) 12px max(16px,var(--sl))}
.top .row{display:flex;align-items:center;gap:12px;max-width:760px;margin:auto}
.code{background:var(--code);color:#111;font-weight:800;font-size:22px;letter-spacing:.02em;line-height:1.15;
 padding:3px 10px;border-radius:3px;font-variant-numeric:tabular-nums}
.top .lbl{font-size:13px;opacity:.75;line-height:1.2}
.logo{height:24px;width:auto;flex:none}
@media (max-width:420px){.logo~.lbl{display:none}}
.cont{margin-left:auto;font-size:15px;font-variant-numeric:tabular-nums;white-space:nowrap}
.seg{display:flex;gap:4px;max-width:760px;margin:10px auto 0}
.seg i{flex:1;height:5px;border-radius:3px;background:#ffffff2e;transition:background .25s}
.seg i.on{background:#7fd69f}

main{max-width:760px;margin:auto;display:grid;gap:22px;
 padding:16px max(16px,var(--sr)) calc(32px + var(--sb)) max(16px,var(--sl))}
h2{font-size:15px;font-weight:600;color:var(--muted);margin:0 0 8px}
.pronto{background:var(--okbg);color:var(--ok);border:1px solid var(--okline);border-radius:10px;padding:12px 14px;font-weight:600}
.pronto span{display:block;font-weight:400;color:var(--ink);margin-top:2px}

.lista{display:grid;gap:8px}
@media (min-width:620px){.lista{grid-template-columns:1fr 1fr}}
.slot{display:flex;align-items:center;gap:12px;width:100%;min-height:62px;padding:10px 14px;text-align:left;
 border:1px solid var(--line);border-radius:10px;background:var(--card)}
.slot:active{background:#f5f7f9}
.slot .n{flex:1;min-width:0;font-weight:600;line-height:1.25}
.slot .s{display:block;font-size:13px;font-weight:400;color:var(--muted);margin-top:2px}
.mk{width:28px;height:28px;flex:none;border-radius:50%;border:2px dashed #a9b5bf;display:grid;place-items:center}
.mk svg{display:none}
.slot.ok{background:var(--okbg);border-color:var(--okline)}
.slot.ok .mk{border:0;background:var(--ok)}.slot.ok .mk svg{display:block}
.slot.ok .s{color:var(--ok)}

.ficha{position:relative;background:#fff}
.ficha img.base{width:100%;display:block}
.ficha img.sig{position:absolute;object-fit:contain;object-position:center bottom}
.recorte{overflow:hidden;border:1px solid var(--line);border-radius:8px;background:#fff}
.recorte .ficha{will-change:transform}
.acoes{display:flex;flex-wrap:wrap;gap:8px;margin-top:10px}
.btn{border:1px solid var(--line);background:var(--card);border-radius:8px;padding:10px 14px;font-weight:600;text-decoration:none;color:var(--ink);display:inline-block}
.semimg{padding:14px;color:var(--muted)}

#viewer{position:fixed;inset:0;z-index:20;background:#1b2530;overflow:auto;overscroll-behavior:contain}
#viewer .vbar{position:sticky;top:0;left:0;display:flex;gap:8px;justify-content:flex-end;z-index:1;
 padding:calc(8px + var(--st)) max(10px,var(--sr)) 8px max(10px,var(--sl));background:#1b2530e6}
#viewer .vbar button{border:0;border-radius:8px;padding:9px 14px;background:#fff;font-weight:600}
#viewer .ficha{margin:0 auto 24px;width:100%;max-width:1000px}
#viewer.zoom .ficha{width:230%;max-width:none}

#pad{position:fixed;inset:0;z-index:30;background:var(--card);display:flex;flex-direction:column;overscroll-behavior:contain}
#pad .bar{display:flex;flex-wrap:wrap;align-items:center;gap:6px;background:var(--ink);color:#fff;
 padding:calc(6px + var(--st)) max(8px,var(--sr)) 6px max(8px,var(--sl))}
#pad .bar b{flex:1;min-width:110px;font-size:16px;padding:0 4px}
#pad .bar button{border:0;border-radius:8px;padding:10px 13px;background:#ffffff1f;color:#fff}
#pad .bar .save{background:#2f9e5b;font-weight:700}
.nomebox{display:flex;gap:6px;background:#f3f5f8;padding:6px max(8px,var(--sr)) 6px max(8px,var(--sl))}
.nomebox input{flex:1;min-width:0;font:inherit;font-size:16px;padding:9px;border:1px solid var(--line);border-radius:8px}
.nomebox button{border:0;border-radius:8px;padding:9px 14px;background:var(--ink);color:#fff;font-weight:600}
#msg{background:var(--warnbg);color:var(--warn);padding:7px 12px;font-size:14px}
.area{flex:1;position:relative;background:var(--card)}
#cv{position:absolute;inset:0;width:100%;height:100%;touch-action:none;cursor:crosshair}
.guia{position:absolute;left:7%;right:7%;bottom:36%;border-bottom:2px dashed #b7c2cc;pointer-events:none}
.guia span{position:absolute;left:0;top:8px;font-size:13px;color:var(--muted);white-space:nowrap}
#gire{position:absolute;top:10px;left:50%;transform:translateX(-50%);background:#fff3cd;color:#664d03;
 padding:6px 12px;border-radius:8px;font-size:13px;pointer-events:none;display:none}
@media (orientation:portrait) and (pointer:coarse){#gire{display:block}}

#toast{position:fixed;left:50%;bottom:calc(18px + var(--sb));transform:translate(-50%,20px);opacity:0;z-index:40;
 background:var(--ink);color:#fff;padding:10px 16px;border-radius:999px;font-size:15px;transition:.25s;pointer-events:none;max-width:calc(100% - 32px)}
#toast.on{opacity:1;transform:translate(-50%,0)}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style>

<header class=top>
 <div class=row><img class=logo src="/logo.svg" alt="" onerror="this.remove()"><span class=code>__FACA__</span><span class=lbl>Ficha da faca<br>coleta de assinaturas</span>
  <span class=cont id=cont></span></div>
 <div class=seg id=seg></div>
</header>

<main>
 <div class=pronto id=pronto hidden>Todas as assinaturas coletadas.
  <span>No Excel: Assinaturas &gt; Finalizar coleta, para gerar o PDF assinado.</span></div>
 <section>
  <h2>Toque no setor para assinar</h2>
  <div class=lista id=lista></div>
 </section>
 <section>
  <h2>Responsáveis na ficha</h2>
  <div class=recorte id=recorte><div class=ficha id=fichaP>
   <img class=base src="/ficha.png?t=__T__" alt="Bloco de responsáveis da ficha" onload="recortar(this)" onerror="semImagem()">
   <div class=sigs></div></div></div>
  <div class=acoes><button class=btn id=verFicha onclick="verFicha()">Ver ficha inteira</button>
   <a class=btn id=pdf href="/ficha.pdf?t=__T__" target=_blank>Abrir PDF</a></div>
 </section>
</main>

<div id=viewer hidden>
 <div class=vbar><button onclick="zoom()" id=zbtn>Ampliar</button><button onclick="fecharFicha()">Fechar</button></div>
 <div class=ficha id=fichaV><img class=base src="/ficha.png?t=__T__" alt="Ficha da faca"><div class=sigs></div></div>
</div>

<div id=pad hidden>
 <div class=bar><button onclick="fechar()">Cancelar</button><b id=titulo></b>
  <button onclick="digitar()">Digitar nome</button><button onclick="limpar()">Limpar</button>
  <button class=save onclick="salvar()">Salvar</button></div>
 <div class=nomebox id=nomebox hidden><input id=nomein placeholder="Nome do gestor" autocomplete=off enterkeyhint=done
  onkeydown="if(event.key=='Enter')aplicarNome()"><button onclick="aplicarNome()">Usar nome</button></div>
 <div id=msg hidden></div>
 <div class=area><canvas id=cv></canvas><div class=guia><span id=legenda></span></div>
  <div id=gire>Gire o celular para assinar deitado</div></div>
</div>

<div id=toast role=status aria-live=polite></div>

<script>
const T="__T__",FACA="__FACA__",SLOTS=__SLOTS__;
const $=id=>document.getElementById(id),cv=$('cv'),x=cv.getContext('2d');
const SRC=window.SRC_DEMO||((k,v)=>`/sig/${k}.png?t=${T}&v=${v}`);
const CHECK='<svg width=16 height=16 viewBox="0 0 16 16" aria-hidden=true><path d="M3 8.5l3.2 3L13 4.8" fill=none stroke="#fff" stroke-width=2.4 stroke-linecap=round stroke-linejoin=round /></svg>';
let done=new Set(),cur=null,box=null,modo='desenho',on=false,pts=[];
const nome=k=>SLOTS.find(s=>s[0]==k)[1];
const aviso=t=>{$('msg').textContent=t;$('msg').hidden=!t};
let tt;const toast=t=>{$('toast').textContent=t;$('toast').classList.add('on');clearTimeout(tt);tt=setTimeout(()=>$('toast').classList.remove('on'),2600)};

function render(){
 const n=done.size,tot=SLOTS.length;
 $('cont').textContent=n+' de '+tot;
 $('seg').innerHTML=SLOTS.map(([k])=>`<i class="${done.has(k)?'on':''}"></i>`).join('');
 $('pronto').hidden=n<tot;
 $('lista').innerHTML=SLOTS.map(([k,nm])=>{const ok=done.has(k);
  return `<button class="slot${ok?' ok':''}" onclick="abrir('${k}')"><span class=mk>${CHECK}</span>
   <span class=n>${nm}<span class=s>${ok?'Assinado, toque para refazer':'Aguardando assinatura'}</span></span></button>`}).join('');
 const v=Date.now(),sig=SLOTS.filter(s=>done.has(s[0])).map(([k,,l,t,w,h])=>
  `<img class=sig alt="" src="${SRC(k,v)}" style="left:${l*100}%;top:${t*100}%;width:${w*100}%;height:${h*100}%">`).join('');
 document.querySelectorAll('.sigs').forEach(e=>e.innerHTML=sig);
}
function status(){if(!cur)fetch('/status?t='+T).then(r=>r.json()).then(l=>{if(l.sort().join()!=[...done].sort().join()){done=new Set(l);render()}}).catch(()=>{})}

// mostra so o bloco RESPONSAVEIS da imagem da ficha
function recortar(img){
 const t=Math.max(0,Math.min(...SLOTS.map(s=>s[3]))-.035),b=Math.min(1,Math.max(...SLOTS.map(s=>s[3]+s[5]))+.03);
 $('recorte').style.aspectRatio=`${img.naturalWidth}/${img.naturalHeight*(b-t)}`;
 $('fichaP').style.transform=`translateY(${-t*100}%)`;
}
function semImagem(){$('recorte').outerHTML='<p class=semimg>Imagem da ficha indisponível. Use o PDF.</p>';$('verFicha').hidden=true}
function verFicha(){$('viewer').hidden=false;document.body.style.overflow='hidden'}
function fecharFicha(){$('viewer').hidden=true;$('viewer').classList.remove('zoom');$('zbtn').textContent='Ampliar';document.body.style.overflow=''}
function zoom(){const z=$('viewer').classList.toggle('zoom');$('zbtn').textContent=z?'Reduzir':'Ampliar'}

function fit(){
 const r=cv.getBoundingClientRect(),d=devicePixelRatio||1;
 cv.width=r.width*d;cv.height=r.height*d;x.setTransform(d,0,0,d,0,0);
 x.lineWidth=2.8;x.lineCap=x.lineJoin='round';x.strokeStyle=x.fillStyle=getComputedStyle(document.body).getPropertyValue('--pen');
 box=null;modo='desenho';
}
addEventListener('resize',()=>{if(cur&&!box)fit()});  // girar antes de assinar reajusta a tela
const P=e=>{const r=cv.getBoundingClientRect();return[e.clientX-r.left,e.clientY-r.top]};
const grow=p=>box=box?[Math.min(box[0],p[0]),Math.min(box[1],p[1]),Math.max(box[2],p[0]),Math.max(box[3],p[1])]:[p[0],p[1],p[0],p[1]];
const mid=(a,b)=>[(a[0]+b[0])/2,(a[1]+b[1])/2];
cv.onpointerdown=e=>{
 if(modo=='digitado')fit();
 on=true;cv.setPointerCapture(e.pointerId);const p=P(e);pts=[p];grow(p);
 x.beginPath();x.arc(p[0],p[1],x.lineWidth/2,0,7);x.fill();
};
cv.onpointermove=e=>{
 if(!on)return;
 for(const ev of (e.getCoalescedEvents?e.getCoalescedEvents():[e])){
  const p=P(ev);pts.push(p);grow(p);const n=pts.length;x.beginPath();
  if(n==2){x.moveTo(...pts[0]);x.lineTo(...mid(pts[0],p))}
  else{const a=pts[n-3],b=pts[n-2];x.moveTo(...mid(a,b));x.quadraticCurveTo(b[0],b[1],...mid(b,p))}  // traco suavizado
  x.stroke();
 }
};
cv.onpointerup=cv.onpointercancel=()=>{if(on&&pts.length>1){const n=pts.length;x.beginPath();x.moveTo(...mid(pts[n-2],pts[n-1]));x.lineTo(...pts[n-1]);x.stroke()}on=false};

async function abrir(k){
 cur=k;$('titulo').textContent=nome(k);aviso('');$('nomebox').hidden=true;
 $('legenda').textContent=`${nome(k)}, faca ${FACA}, ${new Date().toLocaleDateString('pt-BR')}`;
 $('pad').hidden=false;document.body.style.overflow='hidden';
 try{await document.documentElement.requestFullscreen();await screen.orientation.lock('landscape')}catch(e){}  // iPhone nao deixa: aparece o aviso
 requestAnimationFrame(fit);
}
function fechar(){
 $('pad').hidden=true;cur=null;document.body.style.overflow='';
 try{screen.orientation.unlock()}catch(e){}
 if(document.fullscreenElement)document.exitFullscreen().catch(()=>{});
}
function limpar(){fit();aviso('')}
function digitar(){$('nomebox').hidden=false;$('nomein').value='';$('nomein').focus()}
function aplicarNome(){
 const n=$('nomein').value.trim();if(!n)return;
 $('nomein').blur();$('nomebox').hidden=true;
 setTimeout(()=>{fit();escreve(n)},350);  // espera o teclado do celular fechar e a tela reajustar
}
function escreve(n){
 const r=cv.getBoundingClientRect(),base=r.height*.64;  // mesma altura da .guia (bottom 36%)
 const f=s=>`${s}px 'Segoe Script','Brush Script MT','Snell Roundhand','Dancing Script',cursive`;
 let fs=Math.min(r.height*.32,84);x.font=f(fs);let w=x.measureText(n).width;
 if(w>r.width*.84){fs*=r.width*.84/w;x.font=f(fs);w=x.measureText(n).width}
 const x0=(r.width-w)/2;x.fillText(n,x0,base-6);
 box=[x0,base-6-fs,x0+w,base-6+fs*.35];modo='digitado';aviso('');
}
async function salvar(){
 if(!box){aviso('Assine (ou digite o nome) antes de salvar.');return}
 const r=cv.getBoundingClientRect(),d=devicePixelRatio||1,m=10;
 const a=Math.max(0,box[0]-m),b=Math.max(0,box[1]-m),c=Math.min(r.width,box[2]+m),e=Math.min(r.height,box[3]+m);
 const o=document.createElement('canvas');o.width=Math.max(1,(c-a)*d);o.height=Math.max(1,(e-b)*d);  // recorta no traco
 o.getContext('2d').drawImage(cv,a*d,b*d,o.width,o.height,0,0,o.width,o.height);
 let res;try{res=await fetch(`/save/${cur}?t=${T}`,{method:'POST',headers:{'X-Modo':modo},body:o.toDataURL('image/png')})}catch(err){}
 if(!res||!res.ok){aviso('Não salvou'+(res?' (erro '+res.status+')':' (sem conexão com o PC)')+'. Tente de novo.');return}
 const k=cur;done.add(k);fechar();render();
 try{navigator.vibrate&&navigator.vibrate(30)}catch(e){}
 toast(`Assinatura de ${nome(k)} salva`);
}
render();status();setInterval(status,5000);
</script></html>"""

QR_PAGE = """<!doctype html><html lang=pt-br><meta charset=utf-8><title>Coleta __FACA__</title>
<style>body{font-family:system-ui,sans-serif;text-align:center;margin:0 0 40px}
header{background:#13293d;color:#fff;padding:16px;margin-bottom:24px}header img{height:32px;display:block;margin:0 auto 8px}
header h1{margin:0;font-size:24px}
#qr{display:inline-block;padding:16px;background:#fff}code{font-size:26px}
a.pc{display:inline-block;margin-top:8px;padding:10px 18px;background:#1f3b57;color:#fff;border-radius:6px;text-decoration:none}</style>
<header><img src="/logo.svg" alt="" onerror="this.remove()"><h1>Coleta de assinaturas, faca __FACA__</h1></header>
<p>Escaneie com o celular (mesma rede) ou digite:</p>
<div id=qr></div><p><code>__URL__</code></p>
<p>Plano B: <a class=pc href="__URL__">Assinar neste PC</a></p>
<p>Depois das assinaturas, no Excel: <b>Assinaturas &gt; Finalizar coleta</b>.</p>
<!-- shortcut: QR via CDN, precisa de internet no PC; sem ela fica so o link -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script>if(window.QRCode)new QRCode(document.getElementById('qr'),{text:"__URL__",width:280,height:280})</script>
</html>"""


class Servidor(ThreadingHTTPServer):
    # no Windows o reuse deixa 2 coletas na mesma porta (o celular cairia na antiga): porta exclusiva
    allow_reuse_address = False

    def server_bind(self):
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def ip_local():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("10.255.255.255", 1))  # nao envia nada, so descobre a interface
        return s.getsockname()[0]
    except OSError:
        return "127.0.0.1"
    finally:
        s.close()


def make_handler(folder, faca, token, url, slots):
    sigdir = os.path.join(folder, "assinaturas")
    os.makedirs(sigdir, exist_ok=True)
    keys = {s[0] for s in slots}

    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):
            pass

        def _send(self, code, body=b"", ctype="text/plain"):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _auth(self):
            return ("t=" + token) in self.path.partition("?")[2].split("&")

        def _file(self, path, ctype):
            if not os.path.exists(path):
                return self._send(404)
            with open(path, "rb") as f:
                self._send(200, f.read(), ctype)

        def do_GET(self):
            if self.path == "/favicon.ico":  # o navegador pede sem token; evita erro no console
                return self._send(204)
            if self.path == "/logo.svg":  # opcional: logo.svg ao lado do server.py (sem arquivo, a pagina esconde)
                return self._file(os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.svg"), "image/svg+xml")
            if not self._auth():
                return self._send(403, b"acesso negado")
            route = self.path.partition("?")[0]
            if route == "/":
                html = (PAGE.replace("__T__", token).replace("__FACA__", faca)
                        .replace("__SLOTS__", json.dumps(slots)))
                return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            if route == "/qr":
                html = QR_PAGE.replace("__FACA__", faca).replace("__URL__", url)
                return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            if route == "/ficha.pdf":
                return self._file(os.path.join(folder, "ficha.pdf"), "application/pdf")
            if route == "/ficha.png":
                return self._file(os.path.join(folder, "ficha.png"), "image/png")
            if route.startswith("/sig/") and route[5:-4] in keys and route.endswith(".png"):
                return self._file(os.path.join(sigdir, route[5:]), "image/png")
            if route == "/status":
                done = [k for k in keys if os.path.exists(os.path.join(sigdir, k + ".png"))]
                return self._send(200, json.dumps(done).encode(), "application/json")
            self._send(404)

        def do_POST(self):
            route = self.path.partition("?")[0]
            key = route.rpartition("/")[2]
            if not (self._auth() and route.startswith("/save/") and key in keys):
                return self._send(403)
            n = int(self.headers.get("Content-Length", 0))
            if n <= 0 or n > MAX_BYTES:
                return self._send(413)
            raw = self.rfile.read(n).decode("ascii", "ignore")
            try:
                png = base64.b64decode(raw.split(",", 1)[1], validate=True)
            except (IndexError, ValueError):
                return self._send(400)
            if not png.startswith(PNG):
                return self._send(400)
            modo = self.headers.get("X-Modo")
            modo = modo if modo in MODOS else MODOS[0]
            with open(os.path.join(sigdir, key + ".png"), "wb") as f:
                f.write(png)
            with open(os.path.join(folder, "log.csv"), "a", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow([datetime.datetime.now().isoformat(timespec="seconds"),
                                        self.client_address[0], key, modo])
            self._send(200, b"ok")

    return H


def main():
    folder = os.path.abspath(sys.argv[1])
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 5055
    faca = os.path.basename(folder)
    # [chave, titulo, esq, topo, larg, alt] - posicoes em fracao da area de impressao
    arq = os.path.join(folder, "slots.json")
    if os.path.exists(arq):
        with open(arq, encoding="cp1252") as f:  # VBA grava em ANSI
            slots = json.load(f)
    else:  # macro antigo nao grava slots.json: usa o layout padrao da ficha
        slots = SLOTS_PADRAO
    token = "".join(secrets.choice(ALFA) for _ in range(6))
    url = f"http://{ip_local()}:{port}/?t={token}"
    srv = Servidor(("0.0.0.0", port), make_handler(folder, faca, token, url, slots))
    with open(os.path.join(folder, "url.txt"), "w") as f:
        f.write(url)

    def watch_stop():  # o macro cria stop.flag quando termina
        stop = os.path.join(folder, "stop.flag")
        while not os.path.exists(stop):
            threading.Event().wait(1)
        srv.shutdown()

    threading.Thread(target=watch_stop, daemon=True).start()
    srv.serve_forever()


if __name__ == "__main__":
    main()
