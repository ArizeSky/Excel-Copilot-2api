import json
import time
import sys
import websocket
import urllib.request

sys.stdout.reconfigure(encoding='utf-8')

import browser_driver as bd

PROMPT = 'NETWORK_TRACE_PROMPT_20260415'
WATCH_SUBSTRINGS = (
    'augloop.office.com',
    'AcquireTokenForAugloop.ashx',
    'login.microsoftonline.com/common/oauth2/v2.0/token',
    'copilot',
    'chat',
)

targets = bd.load_targets()
taskpane = bd.pick_taskpane_target(targets)
client = bd.CDPClient(taskpane['webSocketDebuggerUrl'])
client.connect()

try:
    client.call('Runtime.enable')
    client.call('Network.enable')

    before = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
    before_nodes = before.get('result', {}).get('value', {}).get('nodes', [])
    before_messages = bd.extract_article_messages(before_nodes)
    print('BEFORE_COUNT', len(before_messages))

    send = client.call('Runtime.evaluate', {'expression': bd.taskpane_send_script(PROMPT), 'returnByValue': True})
    print('SEND_RESULT', json.dumps(send, ensure_ascii=False))

    deadline = time.time() + 30
    body_cmds = {}
    next_id = 1000
    while time.time() < deadline:
        raw = client.ws.recv()
        data = json.loads(raw)
        method = data.get('method')
        params = data.get('params', {})

        if method == 'Network.requestWillBeSent':
            req = params.get('request', {})
            url = req.get('url', '')
            if any(k in url for k in WATCH_SUBSTRINGS):
                print('REQ', url)
                if req.get('postData'):
                    print('POSTDATA', req.get('postData', '')[:5000])

        elif method == 'Network.responseReceived':
            resp = params.get('response', {})
            url = resp.get('url', '')
            req_id = params.get('requestId')
            if any(k in url for k in WATCH_SUBSTRINGS):
                print('RESP', resp.get('status'), url)
                next_id += 1
                body_cmds[next_id] = url
                client.ws.send(json.dumps({'id': next_id, 'method': 'Network.getResponseBody', 'params': {'requestId': req_id}}))

        elif data.get('id') in body_cmds:
            url = body_cmds.pop(data['id'])
            print('BODY_FOR', url)
            print(data.get('result', {}).get('body', '')[:8000])

        current = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
        nodes = current.get('result', {}).get('value', {}).get('nodes', [])
        messages = bd.extract_article_messages(nodes)
        candidate = bd.pick_response_for_prompt(before_messages, messages, PROMPT)
        if candidate:
            print('FINAL_CANDIDATE', json.dumps(candidate, ensure_ascii=False))
            break
finally:
    client.close()
