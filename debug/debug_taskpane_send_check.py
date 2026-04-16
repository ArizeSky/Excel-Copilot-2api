import json
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

import browser_driver as bd

prompt = 'ONLY_APPLE_20260415'
targets = bd.load_targets()
taskpane = bd.pick_taskpane_target(targets)
client = bd.CDPClient(taskpane['webSocketDebuggerUrl'])
client.connect()

try:
    client.call('Runtime.enable')

    before = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
    before_nodes = before.get('result', {}).get('value', {}).get('nodes', [])
    before_msgs = bd.extract_article_messages(before_nodes)
    print('BEFORE_COUNT', len(before_msgs))
    print('BEFORE_LAST2', json.dumps(before_msgs[-2:], ensure_ascii=False))

    send = client.call('Runtime.evaluate', {'expression': bd.taskpane_send_script(prompt), 'returnByValue': True})
    print('SEND_RESULT', json.dumps(send, ensure_ascii=False))

    time.sleep(2)

    after = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
    after_nodes = after.get('result', {}).get('value', {}).get('nodes', [])
    after_msgs = bd.extract_article_messages(after_nodes)
    print('AFTER_COUNT', len(after_msgs))
    print('AFTER_LAST4', json.dumps(after_msgs[-4:], ensure_ascii=False))

    textbox_expr = r'''(() => {
      const n = document.querySelector('[role="textbox"], textarea, [contenteditable="true"], input[type="text"]');
      return n ? ((n.innerText || n.textContent || n.value || '').trim()) : null;
    })()'''
    textbox_state = client.call('Runtime.evaluate', {'expression': textbox_expr, 'returnByValue': True})
    print('TEXTBOX_AFTER', json.dumps(textbox_state.get('result', {}), ensure_ascii=False))
finally:
    client.close()
