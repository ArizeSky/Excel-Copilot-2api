import json
import urllib.request
import websocket
import sys

sys.stdout.reconfigure(encoding='utf-8')

targets = json.load(urllib.request.urlopen('http://127.0.0.1:9222/json'))
pane = [t for t in targets if t.get('type') == 'iframe' and 'taskpane.html' in t.get('url', '')][0]
ws = websocket.create_connection(pane['webSocketDebuggerUrl'], timeout=30, origin='http://127.0.0.1:9222')

try:
    ws.send(json.dumps({'id': 1, 'method': 'Runtime.enable', 'params': {}}))
    while True:
        data = json.loads(ws.recv())
        if data.get('id') == 1:
            break

    expr = r'''(() => {
      const selector = '[data-testid], [role="article"], .message, .response, .markdown, textarea, [contenteditable="true"], button';
      const nodes = [...document.querySelectorAll(selector)].slice(0, 200);
      return nodes.map((n, i) => ({
        i,
        tag: n.tagName,
        role: n.getAttribute('role'),
        testid: n.getAttribute('data-testid'),
        cls: (n.className || '').toString().slice(0, 200),
        text: (n.innerText || n.textContent || '').trim().slice(0, 300)
      }));
    })()'''

    ws.send(json.dumps({'id': 2, 'method': 'Runtime.evaluate', 'params': {'expression': expr, 'returnByValue': True}}))
    while True:
        data = json.loads(ws.recv())
        if data.get('id') == 2:
            value = data.get('result', {}).get('result', {}).get('value', [])
            print(json.dumps(value[:120], ensure_ascii=False, indent=2))
            break
finally:
    ws.close()
