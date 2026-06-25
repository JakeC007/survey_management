"""
examine_app.py
----------------
Local web front end for examining an individual survey participant.

Run it:
    python examine_app.py
then open the URL it prints (defaults to http://127.0.0.1:8765).

You type a name (first, or first + last), a response/RID (R_xxx), or an email.
It finds the person in participant_tracker_auto, confirms whether they have a
completed survey, tells you their payment status (PAY / HOLD / FRAUD / PAID)
and why, and shows their answers plus key quality metadata.

No third-party web framework needed: this uses only the Python standard
library plus openpyxl (already in the project's requirements). Data is loaded
once at startup from examine_data.py.
"""

from __future__ import annotations

import json
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

import examine_data

HOST = "127.0.0.1"
PORT = 8765

STORE = None  # populated in main()


# --------------------------------------------------------------------------- #
# The single-page UI
# --------------------------------------------------------------------------- #
INDEX_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Examine participant</title>
<style>
  :root{
    --bg:#0f1115; --panel:#171a21; --panel2:#1d212b; --line:#2a2f3a;
    --text:#e7e9ee; --muted:#9aa3b2; --accent:#5b9dff;
    --pay:#22c55e; --hold:#f59e0b; --fraud:#ef4444; --paid:#3b82f6; --na:#6b7280;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
       background:var(--bg);color:var(--text);line-height:1.45}
  header{padding:22px 28px 14px;border-bottom:1px solid var(--line)}
  h1{font-size:19px;margin:0 0 3px}
  .sub{color:var(--muted);font-size:13px}
  .wrap{max-width:980px;margin:0 auto;padding:22px 28px 60px}
  .searchbar{display:flex;gap:10px;margin-bottom:6px}
  input[type=text]{flex:1;background:var(--panel);border:1px solid var(--line);color:var(--text);
       padding:13px 14px;border-radius:10px;font-size:15px;outline:none}
  input[type=text]:focus{border-color:var(--accent)}
  button{background:var(--accent);color:#07101f;border:0;border-radius:10px;padding:0 20px;
       font-size:15px;font-weight:600;cursor:pointer}
  button:hover{filter:brightness(1.08)}
  .hint{color:var(--muted);font-size:12px;margin:8px 2px 18px}
  .candidates{display:flex;flex-direction:column;gap:8px;margin-top:8px}
  .cand{display:flex;align-items:center;gap:12px;background:var(--panel);border:1px solid var(--line);
       border-radius:10px;padding:11px 14px;cursor:pointer}
  .cand:hover{border-color:var(--accent)}
  .cand .who{flex:1;min-width:0}
  .cand .who .nm{font-weight:600}
  .cand .who .meta{color:var(--muted);font-size:12.5px;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .badge{font-size:11px;font-weight:700;letter-spacing:.04em;padding:4px 9px;border-radius:999px;white-space:nowrap}
  .b-PAY{background:rgba(34,197,94,.15);color:var(--pay)}
  .b-HOLD{background:rgba(245,158,11,.15);color:var(--hold)}
  .b-FRAUD{background:rgba(239,68,68,.15);color:var(--fraud)}
  .b-PAID{background:rgba(59,130,246,.15);color:var(--paid)}
  .b-NA,.b-NOTEVALUATED{background:rgba(107,114,128,.18);color:#b6bdca}
  .pill{font-size:11px;color:var(--muted);border:1px solid var(--line);border-radius:999px;padding:3px 8px}
  .pill.done{color:var(--pay);border-color:rgba(34,197,94,.4)}
  .pill.notdone{color:var(--hold);border-color:rgba(245,158,11,.4)}

  .detail{margin-top:10px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 20px;margin-bottom:16px}
  .statusrow{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
  .bigbadge{font-size:15px;font-weight:800;letter-spacing:.05em;padding:8px 16px;border-radius:10px}
  .why{margin-top:12px;color:#d9dde6;font-size:14px;background:var(--panel2);border-left:3px solid var(--line);
       padding:10px 14px;border-radius:0 8px 8px 0}
  .why.fraud{border-color:var(--fraud)} .why.hold{border-color:var(--hold)}
  .why.pay{border-color:var(--pay)} .why.paid{border-color:var(--paid)}
  .name{font-size:20px;font-weight:700}
  .idline{color:var(--muted);font-size:13px;margin-top:3px}
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px 18px;margin-top:6px}
  .kv .k{color:var(--muted);font-size:11.5px;text-transform:uppercase;letter-spacing:.04em}
  .kv .v{font-size:14px;margin-top:1px;word-break:break-word}
  h2{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin:0 0 12px}
  .qa{border-top:1px solid var(--line);padding:11px 0}
  .qa:first-of-type{border-top:0}
  .qa .q{font-size:13.5px;color:var(--muted)}
  .qa .a{font-size:15px;margin-top:3px;white-space:pre-wrap}
  .qa.other .q::after{content:" (free text)";color:var(--accent);font-size:11px}
  .empty{color:var(--muted);font-style:italic}
  .notice{background:rgba(245,158,11,.10);border:1px solid rgba(245,158,11,.35);color:#f6d08a;
       border-radius:10px;padding:12px 14px;font-size:14px}
  .back{background:none;color:var(--accent);padding:0;font-weight:600;font-size:14px;margin-bottom:14px}
  .back:hover{text-decoration:underline}
  .err{color:#f6a6a6}
  code{background:var(--panel2);padding:1px 6px;border-radius:5px;font-size:12.5px}
</style>
</head>
<body>
<header>
  <h1>Examine a participant</h1>
  <div class="sub">Search the participant tracker, confirm survey completion, and read their answers with payment status.</div>
</header>
<div class="wrap">
  <div class="searchbar">
    <input id="q" type="text" autofocus placeholder="Name (first, or first + last), R_… response ID, or email"
           onkeydown="if(event.key==='Enter')doSearch()">
    <button onclick="doSearch()">Search</button>
  </div>
  <div class="hint">Examples: <code>Harper Patrick</code> &nbsp;·&nbsp; <code>Eric</code> (shows every match) &nbsp;·&nbsp; <code>R_3Q7zMYZjRHFPI2d</code> &nbsp;·&nbsp; <code>name@email.com</code></div>
  <div id="out"></div>
</div>

<script>
const out = document.getElementById('q');
const dst = document.getElementById('out');

function esc(s){return (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function badgeClass(s){return 'b-'+String(s).replace(/[^A-Z]/g,'');}

async function doSearch(){
  const q = document.getElementById('q').value.trim();
  if(!q){dst.innerHTML='';return;}
  dst.innerHTML = '<div class="hint">Searching…</div>';
  let res;
  try{ res = await (await fetch('/api/search?q='+encodeURIComponent(q))).json(); }
  catch(e){ dst.innerHTML='<div class="err">Request failed: '+esc(e)+'</div>'; return; }
  renderCandidates(res);
}

function renderCandidates(res){
  const c = res.candidates||[];
  if(c.length===0){
    dst.innerHTML = '<div class="card"><div class="empty">No one found for “'+esc(res.query)+'”. '
      + 'Try a first name only, a partial name, the R_ id, or an email.</div></div>';
    return;
  }
  if(c.length===1){ openDetail(c[0].response_id); return; }
  let h = '<div class="hint">'+c.length+' matches for “'+esc(res.query)+'” ('+esc(res.kind)+'). Pick one:</div><div class="candidates">';
  for(const x of c){
    const done = x.completed
      ? '<span class="pill done">survey complete</span>'
      : '<span class="pill notdone">no completed survey</span>';
    h += '<div class="cand" onclick="openDetail(\''+x.response_id+'\')">'
       + '<div class="who"><div class="nm">'+esc(x.child_name||'(no child name)')+'</div>'
       + '<div class="meta">parent: '+esc(x.parent_name||'—')+' · '+esc(x.delivery_email||'—')
       + ' · matched on '+esc(x.matched_on)+' · '+esc(x.response_id)+'</div></div>'
       + done
       + '<span class="badge '+badgeClass(x.status)+'">'+esc(x.status)+'</span></div>';
  }
  h += '</div>';
  dst.innerHTML = h;
}

async function openDetail(rid){
  dst.innerHTML = '<div class="hint">Loading…</div>';
  let d;
  try{ d = await (await fetch('/api/detail?rid='+encodeURIComponent(rid))).json(); }
  catch(e){ dst.innerHTML='<div class="err">Request failed: '+esc(e)+'</div>'; return; }
  if(d.error){ dst.innerHTML='<div class="card err">'+esc(d.error)+'</div>'; return; }
  renderDetail(d);
}

function renderDetail(d){
  const p = d.person;
  const st = d.status;
  const lc = st==='FRAUD'?'fraud':st==='HOLD'?'hold':st==='PAID'?'paid':st==='PAY'?'pay':'';
  let h = '<div class="detail"><button class="back" onclick="doSearch()">← back to results</button>';

  // status + person
  h += '<div class="card">';
  h += '<div class="statusrow"><span class="bigbadge '+badgeClass(st)+'">'+esc(st)+'</span>'
     + '<div><div class="name">'+esc(p.child_name||'(no child name)')+'</div>'
     + '<div class="idline">parent/guardian: '+esc(p.parent_name||'—')+' &nbsp;·&nbsp; '+esc(p.delivery_email||'—')
     + ' &nbsp;·&nbsp; grade '+esc(p.grade||'—')+' &nbsp;·&nbsp; '+esc(p.response_id)+'</div></div></div>';
  h += '<div class="why '+lc+'"><b>Why '+esc(st)+':</b> '+esc(d.reason||'—')+'</div>';
  h += '</div>';

  // completion notice
  if(!d.completed){
    h += '<div class="card"><div class="notice">No completed survey on file for this person, so there are no answers to show. '
       + 'The status above reflects what the tracker knows so far.</div></div>';
  }

  // metadata
  if(d.metadata && d.metadata.length){
    h += '<div class="card"><h2>Key metadata &amp; quality signals</h2><div class="grid">';
    for(const m of d.metadata){
      h += '<div class="kv"><div class="k">'+esc(m.label)+'</div><div class="v">'+esc(m.value)+'</div></div>';
    }
    h += '</div></div>';
  }

  // answers
  if(d.completed){
    h += '<div class="card"><h2>Survey answers ('+ (d.answers?d.answers.length:0) +')</h2>';
    if(d.answers && d.answers.length){
      for(const a of d.answers){
        h += '<div class="qa'+(a.other?' other':'')+'"><div class="q">'+esc(a.label)+'</div>'
           + '<div class="a">'+esc(a.value)+'</div></div>';
      }
    }else{
      h += '<div class="empty">Marked complete, but no answer fields were found in the export ('+esc(d.export_name||'?')+').</div>';
    }
    h += '</div>';
  }

  h += '<div class="hint">Answers sourced from <code>'+esc(d.export_name||'?')+'</code>.</div></div>';
  dst.innerHTML = h;
}
</script>
</body>
</html>
"""


# --------------------------------------------------------------------------- #
# HTTP handler
# --------------------------------------------------------------------------- #
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # quiet

    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj), "application/json; charset=utf-8")

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if path in ("/", "/index.html"):
            self._send(200, INDEX_HTML, "text/html; charset=utf-8")
            return

        if path == "/api/search":
            q = (qs.get("q") or [""])[0]
            try:
                self._json(STORE.search(q))
            except Exception as e:  # noqa: BLE001
                self._json({"error": f"search failed: {e}"}, 500)
            return

        if path == "/api/detail":
            rid = (qs.get("rid") or [""])[0]
            try:
                d = STORE.detail(rid)
                if d is None:
                    self._json({"error": f"No participant found for {rid!r}."}, 404)
                else:
                    self._json(d)
            except Exception as e:  # noqa: BLE001
                self._json({"error": f"detail failed: {e}"}, 500)
            return

        self._send(404, "not found", "text/plain; charset=utf-8")


# --------------------------------------------------------------------------- #
# Entry point
# --------------------------------------------------------------------------- #
def main():
    global STORE
    print("Loading data (tracker, payments, blacklist, latest survey export)…")
    STORE = examine_data.load_all()
    print(f"  tracker rows : {len(STORE.tracker)}")
    print(f"  payment rows : {len(STORE.pay_by_cid)}")
    print(f"  blacklist    : {len(STORE.bl_rows)}")
    print(f"  answers from : {STORE.export_name}  ({len(STORE.ans_by_cid)} responses)")

    port = PORT
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass

    url = f"http://{HOST}:{port}"
    server = ThreadingHTTPServer((HOST, port), Handler)
    print(f"\nReady. Open {url} in your browser.  (Ctrl+C to stop.)")
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
        server.shutdown()


if __name__ == "__main__":
    main()
