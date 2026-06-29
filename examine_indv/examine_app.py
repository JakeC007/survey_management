"""
examine_app.py
----------------
Local web front end for examining an individual survey participant.

Run it:
    python examine_app.py
then open the URL it prints (defaults to http://127.0.0.1:8765).

Type a name (first, or first + last), a response/RID (R_xxx), or an email.
It finds the person, confirms whether they have a completed survey, tells you
their status (PAY / HOLD / FRAUD / PAID) and why, and shows their answers plus
key quality metadata.

The detail view has a "Mark as fraud" button. It asks for confirmation, then
writes the person into fraud_blacklist.csv (the source the payment pipeline
reads), payment_tracker.xlsx, and payment_report_unpaid.xlsx, so the change is
durable and the dashboard picks it up. See examine_write.py.

Standard library only, plus openpyxl (already in the project requirements).
"""

from __future__ import annotations

import json
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

import examine_data
import examine_write
import examine_review

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
       font-size:15px;font-weight:600;cursor:pointer;height:46px}
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
  .spacer{flex:1}
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
  .back{background:none;color:var(--accent);padding:0;font-weight:600;font-size:14px;margin-bottom:14px;height:auto}
  .back:hover{text-decoration:underline;filter:none}
  .err{color:#f6a6a6}
  code{background:var(--panel2);padding:1px 6px;border-radius:5px;font-size:12.5px}

  /* danger button */
  .btn-danger{background:rgba(239,68,68,.14);color:#ff8b8b;border:1px solid rgba(239,68,68,.5);
       height:38px;padding:0 16px;border-radius:9px;font-size:13.5px;font-weight:600}
  .btn-danger:hover{background:rgba(239,68,68,.22);filter:none}
  .already{color:var(--fraud);font-size:13px;font-weight:600}

  /* modal */
  .overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);display:flex;align-items:center;
       justify-content:center;z-index:50}
  .modal{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:22px 24px;
       width:min(520px,92vw);box-shadow:0 20px 60px rgba(0,0,0,.5)}
  .modal h3{margin:0 0 8px;font-size:17px}
  .modal p{color:var(--muted);font-size:13.5px;margin:0 0 14px}
  .modal textarea{width:100%;min-height:70px;background:var(--panel2);border:1px solid var(--line);
       color:var(--text);border-radius:9px;padding:10px 12px;font-size:14px;resize:vertical;font-family:inherit}
  .modal .what{background:var(--panel2);border-radius:9px;padding:10px 12px;margin:12px 0 4px;font-size:12.5px;color:var(--muted)}
  .modal .what b{color:var(--text)}
  .modal-actions{display:flex;justify-content:flex-end;gap:10px;margin-top:18px}
  .btn-ghost{background:none;border:1px solid var(--line);color:var(--text);height:42px;border-radius:9px}
  .btn-ghost:hover{border-color:var(--accent);filter:none}
  .btn-confirm{background:var(--fraud);color:#fff;height:42px;border-radius:9px}
  .toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:var(--panel2);
       border:1px solid var(--line);border-radius:10px;padding:12px 16px;font-size:13.5px;z-index:60;
       max-width:90vw;box-shadow:0 10px 30px rgba(0,0,0,.4)}
  .toast.ok{border-color:rgba(34,197,94,.5)} .toast.bad{border-color:rgba(239,68,68,.5)}
  .toast ul{margin:6px 0 0;padding-left:18px} .toast li{margin:2px 0}
</style>
</head>
<body>
<header>
  <h1>Examine a participant</h1>
  <div class="sub">Search the participant tracker, confirm survey completion, and read their answers with payment status.</div>
  <div style="margin-top:10px"><a href="/review" style="color:var(--accent);font-size:13.5px;font-weight:600;text-decoration:none">→ Review flagged participants (Back / Next queue)</a></div>
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
<div id="modal-root"></div>
<div id="toast-root"></div>

<script>
function esc(s){return (s==null?'':String(s)).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function badgeClass(s){return 'b-'+String(s).replace(/[^A-Z]/g,'');}
let CURRENT=null; // current detail data

async function doSearch(){
  const q = document.getElementById('q').value.trim();
  if(!q){document.getElementById('out').innerHTML='';return;}
  document.getElementById('out').innerHTML = '<div class="hint">Searching…</div>';
  let res;
  try{ res = await (await fetch('/api/search?q='+encodeURIComponent(q))).json(); }
  catch(e){ document.getElementById('out').innerHTML='<div class="err">Request failed: '+esc(e)+'</div>'; return; }
  renderCandidates(res);
}

function renderCandidates(res){
  const dst = document.getElementById('out');
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
  document.getElementById('out').innerHTML = '<div class="hint">Loading…</div>';
  let d;
  try{ d = await (await fetch('/api/detail?rid='+encodeURIComponent(rid))).json(); }
  catch(e){ document.getElementById('out').innerHTML='<div class="err">Request failed: '+esc(e)+'</div>'; return; }
  if(d.error){ document.getElementById('out').innerHTML='<div class="card err">'+esc(d.error)+'</div>'; return; }
  CURRENT=d; renderDetail(d);
}

function renderDetail(d){
  const dst = document.getElementById('out');
  const p = d.person;
  const st = d.status;
  const lc = st==='FRAUD'?'fraud':st==='HOLD'?'hold':st==='PAID'?'paid':st==='PAY'?'pay':'';
  let h = '<div class="detail"><button class="back" onclick="doSearch()">← back to results</button>';

  h += '<div class="card"><div class="statusrow">'
     + '<span class="bigbadge '+badgeClass(st)+'">'+esc(st)+'</span>'
     + '<div><div class="name">'+esc(p.child_name||'(no child name)')+'</div>'
     + '<div class="idline">parent/guardian: '+esc(p.parent_name||'—')+' &nbsp;·&nbsp; '+esc(p.delivery_email||'—')
     + ' &nbsp;·&nbsp; grade '+esc(p.grade||'—')+' &nbsp;·&nbsp; '+esc(p.response_id)+'</div></div>'
     + '<div class="spacer"></div>';
  if(st==='FRAUD'){
    h += '<span class="already">✓ marked fraud</span>';
  }else{
    h += '<button class="btn-danger" onclick="openFraudModal()">Mark as fraud</button>';
  }
  h += '</div>';
  h += '<div class="why '+lc+'"><b>Why '+esc(st)+':</b> '+esc(d.reason||'—')+'</div></div>';

  if(!d.completed){
    h += '<div class="card"><div class="notice">No completed survey on file for this person, so there are no answers to show. '
       + 'The status above reflects what the tracker knows so far.</div></div>';
  }

  if(d.metadata && d.metadata.length){
    h += '<div class="card"><h2>Key metadata &amp; quality signals</h2><div class="grid">';
    for(const m of d.metadata){
      h += '<div class="kv"><div class="k">'+esc(m.label)+'</div><div class="v">'+esc(m.value)+'</div></div>';
    }
    h += '</div></div>';
  }

  if(d.completed && d.featured_text && d.featured_text.length){
    h += '<div class="card"><h2>Open-text responses</h2>';
    for(const a of d.featured_text){
      h += '<div class="qa other"><div class="q">'+esc(a.label)+'</div>'
         + '<div class="a">'+esc(a.value)+'</div></div>';
    }
    h += '</div>';
  }

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

/* ---- Mark-as-fraud confirmation modal ---- */
function openFraudModal(){
  if(!CURRENT) return;
  const p = CURRENT.person;
  const root = document.getElementById('modal-root');
  root.innerHTML =
    '<div class="overlay" onclick="if(event.target===this)closeModal()">'
    + '<div class="modal">'
    + '<h3>Mark this person as fraud?</h3>'
    + '<p>'+esc(p.child_name||p.response_id)+' ('+esc(p.delivery_email||'no email')+') · '+esc(p.response_id)+'</p>'
    + '<label class="k" style="color:var(--muted);font-size:12px">Reason (optional)</label>'
    + '<textarea id="fraud-reason" placeholder="e.g. duplicate submission / fabricated identity / manual review"></textarea>'
    + '<div class="what">This will write to: <b>fraud_blacklist.csv</b> (the source the payment pipeline reads), '
    + '<b>payment_tracker.xlsx</b> (fraud=yes), and <b>payment_report_unpaid.xlsx</b> (moved to the Fraud sheet). '
    + 'The dashboard will reflect it. This person will never be paid.</div>'
    + '<div class="modal-actions">'
    + '<button class="btn-ghost" onclick="closeModal()">Cancel</button>'
    + '<button class="btn-confirm" id="fraud-go" onclick="confirmFraud()">Yes, mark as fraud</button>'
    + '</div></div></div>';
}
function closeModal(){ document.getElementById('modal-root').innerHTML=''; }

async function confirmFraud(){
  const rid = CURRENT.person.response_id;
  const reason = (document.getElementById('fraud-reason').value||'').trim();
  const go = document.getElementById('fraud-go');
  go.disabled = true; go.textContent = 'Saving…';
  let res;
  try{
    res = await (await fetch('/api/mark_fraud',{
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({rid, reason})
    })).json();
  }catch(e){ closeModal(); toast(false,'Request failed: '+esc(e)); return; }
  closeModal();
  if(res.error){ toast(false, esc(res.error)); return; }
  const r = res.result || {};
  const changed = (r.changed||[]).concat(r.skipped||[]);
  let msg = r.ok ? '<b>Marked as fraud.</b>' : '<b>Marked with problems.</b>';
  if(changed.length) msg += '<ul>'+changed.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>';
  if((r.errors||[]).length) msg += '<ul>'+r.errors.map(x=>'<li class="err">'+esc(x)+'</li>').join('')+'</ul>';
  toast(r.ok, msg);
  if(res.detail){ CURRENT=res.detail; renderDetail(res.detail); }
}

function toast(ok, html){
  const root = document.getElementById('toast-root');
  root.innerHTML = '<div class="toast '+(ok?'ok':'bad')+'">'+html+'</div>';
  setTimeout(()=>{ if(root.firstChild) root.innerHTML=''; }, 9000);
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

        if path == "/review":
            self._send(200, examine_review.REVIEW_HTML, "text/html; charset=utf-8")
            return

        if path == "/api/review_queue":
            include_paid = (qs.get("include_paid") or ["1"])[0] != "0"
            try:
                self._json(examine_review.build_queue(STORE, include_paid=include_paid))
            except Exception as e:  # noqa: BLE001
                self._json({"error": f"review_queue failed: {e}"}, 500)
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
                self._json(d if d is not None else {"error": f"No participant found for {rid!r}."},
                           200 if d is not None else 404)
            except Exception as e:  # noqa: BLE001
                self._json({"error": f"detail failed: {e}"}, 500)
            return

        self._send(404, "not found", "text/plain; charset=utf-8")

    def do_POST(self):
        global STORE
        parsed = urlparse(self.path)
        if parsed.path not in ("/api/mark_fraud", "/api/clear_review"):
            self._send(404, "not found", "text/plain; charset=utf-8")
            return
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length) or b"{}")
        except Exception as e:  # noqa: BLE001
            self._json({"error": f"bad request: {e}"}, 400)
            return

        rid = (body.get("rid") or "").strip()
        if not rid:
            self._json({"error": "missing rid"}, 400)
            return

        if parsed.path == "/api/clear_review":
            note = (body.get("note") or "").strip()
            try:
                result = examine_review.clear_item(STORE, rid, note)
                self._json({"result": result})
            except Exception as e:  # noqa: BLE001
                self._json({"error": f"clear_review failed: {e}"}, 500)
            return

        # /api/mark_fraud
        reason = (body.get("reason") or "").strip()
        try:
            result = examine_write.mark_fraud(STORE, rid, reason)
            # reload data so the (now fraud) status is reflected everywhere
            STORE = examine_data.load_all()
            detail = STORE.detail(rid)
            self._json({"result": result, "detail": detail})
        except Exception as e:  # noqa: BLE001
            self._json({"error": f"mark_fraud failed: {e}"}, 500)


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
