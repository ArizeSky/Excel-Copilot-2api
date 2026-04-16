import json
import urllib.request
import websocket
import time


def main():
    targets = json.load(urllib.request.urlopen('http://127.0.0.1:9222/json'))
    target = [t for t in targets if t.get('type') == 'iframe' and 'excel.officeapps.live.com' in t.get('url', '')][0]
    ws = websocket.create_connection(target['webSocketDebuggerUrl'], timeout=10, origin='http://127.0.0.1:9222')
    try:
        ws.send(json.dumps({'id': 1, 'method': 'Page.enable', 'params': {}}))
        ws.send(json.dumps({'id': 2, 'method': 'Network.enable', 'params': {}}))
        ws.send(json.dumps({'id': 3, 'method': 'Page.reload', 'params': {'ignoreCache': True}}))
        deadline = time.time() + 30
        pending_body_request_id = None
        saw = False
        while time.time() < deadline:
            data = json.loads(ws.recv())
            method = data.get('method')
            params = data.get('params', {})

            if method == 'Network.requestWillBeSent':
                req = params.get('request', {})
                url = req.get('url', '')
                if 'AcquireTokenForAugloop.ashx' in url:
                    saw = True
                    print('REQUEST URL:', url)
                    print('REQUEST POSTDATA:', req.get('postData', '')[:12000])

            if method == 'Network.responseReceived':
                resp = params.get('response', {})
                url = resp.get('url', '')
                if 'AcquireTokenForAugloop.ashx' in url:
                    print('RESPONSE STATUS:', resp.get('status'))
                    pending_body_request_id = params.get('requestId')
                    ws.send(json.dumps({'id': 99, 'method': 'Network.getResponseBody', 'params': {'requestId': pending_body_request_id}}))

            if data.get('id') == 99:
                print('RESPONSE BODY:', data.get('result', {}).get('body', '')[:12000])
                return

        print('SAW REQUEST:', saw)
    finally:
        ws.close()


if __name__ == '__main__':
    main()
