import json
import time
import sys

sys.stdout.reconfigure(encoding='utf-8')

import browser_driver as bd

prompt = 'JUST_APPLE_FINAL_20260415'
targets = bd.load_targets()
taskpane = bd.pick_taskpane_target(targets)
client = bd.CDPClient(taskpane['webSocketDebuggerUrl'])
client.connect()

try:
    client.call('Runtime.enable')
    before = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
    before_nodes = before.get('result', {}).get('value', {}).get('nodes', [])
    before_messages = bd.extract_article_messages(before_nodes)
    print('BEFORE_COUNT', len(before_messages))

    send = client.call('Runtime.evaluate', {'expression': bd.taskpane_send_script(prompt), 'returnByValue': True})
    print('SEND_RESULT', json.dumps(send, ensure_ascii=False))

    for i in range(1, 21):
        time.sleep(1)
        after = client.call('Runtime.evaluate', {'expression': bd.taskpane_read_script(), 'returnByValue': True})
        after_nodes = after.get('result', {}).get('value', {}).get('nodes', [])
        after_messages = bd.extract_article_messages(after_nodes)
        candidate = bd.pick_response_for_prompt(before_messages, after_messages, prompt)
        print('TICK', i, 'COUNT', len(after_messages), 'CANDIDATE', json.dumps(candidate, ensure_ascii=False))
        print('LAST2', json.dumps(after_messages[-2:], ensure_ascii=False))
finally:
    client.close()
