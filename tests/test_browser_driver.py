import time
import unittest
import importlib.util
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parent.parent / "browser_driver.py"
spec = importlib.util.spec_from_file_location("browser_driver", MODULE_PATH)
browser_driver = importlib.util.module_from_spec(spec)
spec.loader.exec_module(browser_driver)


class ShellTargetTests(unittest.TestCase):
    def test_pick_shell_target_prefers_excel_cloud_page(self):
        targets = [
            {"type": "page", "url": "https://example.com", "webSocketDebuggerUrl": "ws://debug/1"},
            {"type": "page", "url": "https://excel.cloud.microsoft/open/onedrive/?docId=abc", "webSocketDebuggerUrl": "ws://debug/2"},
        ]
        best = browser_driver.pick_shell_target(targets)
        self.assertEqual(best["webSocketDebuggerUrl"], "ws://debug/2")


class TaskpaneTargetTests(unittest.TestCase):
    def test_pick_taskpane_target_prefers_taskpane_html_iframe(self):
        targets = [
            {"type": "iframe", "url": "https://excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx?wopisrc=x", "webSocketDebuggerUrl": "ws://debug/editor"},
            {"type": "iframe", "url": "https://fa000000124.officeapps.live.com/mro1cdnstorage/foo/taskpane.html?preload=1", "webSocketDebuggerUrl": "ws://debug/taskpane"},
        ]
        best = browser_driver.pick_taskpane_target(targets)
        self.assertEqual(best["webSocketDebuggerUrl"], "ws://debug/taskpane")

    def test_pick_taskpane_target_raises_clear_error(self):
        with self.assertRaisesRegex(RuntimeError, "Copilot taskpane iframe not found"):
            browser_driver.pick_taskpane_target([])


class TaskpaneGlobalTests(unittest.TestCase):
    def test_find_interesting_globals_filters_expected_names(self):
        keys = ["OfficeFirstPartyAuth", "Excel", "document", "location"]
        filtered = browser_driver.find_interesting_globals(keys)
        self.assertEqual(filtered, ["OfficeFirstPartyAuth", "Excel"])


class CDPClientTests(unittest.TestCase):
    def test_call_times_out_when_matching_response_never_arrives(self):
        class FakeWebSocket:
            def __init__(self):
                self.messages = iter([
                    '{"method":"Runtime.consoleAPICalled","params":{}}',
                    '{"method":"Runtime.executionContextCreated","params":{}}',
                ])
                self.sent = []

            def send(self, payload):
                self.sent.append(payload)

            def recv(self):
                try:
                    return next(self.messages)
                except StopIteration:
                    time.sleep(0.001)
                    return '{"method":"Runtime.consoleAPICalled","params":{}}'

        client = browser_driver.CDPClient("ws://debug")
        client.ws = FakeWebSocket()

        with self.assertRaises(TimeoutError):
            client.call("Runtime.evaluate", timeout=0.01)


class TextDeltaTests(unittest.TestCase):
    def test_new_text_delta_returns_suffix_only(self):
        suffix = browser_driver.new_text_delta("hello world", "hello")
        self.assertEqual(suffix, " world")

    def test_new_text_delta_avoids_repeating_common_prefix_after_rewrite(self):
        suffix = browser_driver.new_text_delta("hello there", "hello world")
        self.assertEqual(suffix, "there")


class ResponseTextTests(unittest.TestCase):
    def test_pick_latest_assistant_text_prefers_last_nonempty_message(self):
        messages = ["", "Hi", "Hello from Copilot"]
        text = browser_driver.pick_latest_assistant_text(messages)
        self.assertEqual(text, "Hello from Copilot")

    def test_extract_article_messages_splits_user_and_copilot(self):
        nodes = [
            {"role": "article", "cls": "fai-UserMessage foo", "text": "You said:\nhello"},
            {"role": "article", "cls": "fai-CopilotMessage bar", "text": "Copilot said:\nworld"},
            {"role": None, "cls": "___nk7 narrator-announcement", "text": "ignore me"},
        ]
        messages = browser_driver.extract_article_messages(nodes)
        self.assertEqual(messages, [
            {"kind": "user", "text": "You said:\nhello"},
            {"kind": "assistant", "text": "Copilot said:\nworld"},
        ])

    def test_extract_article_messages_prefers_raw_markdown_for_assistant(self):
        nodes = [
            {"role": "article", "cls": "fai-CopilotMessage bar", "text": "Copilot said:\n标题\n一\n二\nprint(123)", "raw": "# 标题\n- 一\n- 二\n```python\nprint(123)\n```"},
        ]
        messages = browser_driver.extract_article_messages(nodes)
        self.assertEqual(messages, [
            {"kind": "assistant", "text": "# 标题\n- 一\n- 二\n```python\nprint(123)\n```"},
        ])

    def test_extract_article_messages_prefers_raw_final_answer_over_reasoning_pollution(self):
        nodes = [
            {
                "role": "article",
                "cls": "fai-CopilotMessage bar",
                "text": "Copilot said:\n正在思考\n搜索资料\n{\"answer\": 1}",
                "raw": "{\"answer\": 1}",
            },
        ]
        messages = browser_driver.extract_article_messages(nodes)
        self.assertEqual(messages, [
            {"kind": "assistant", "text": "{\"answer\": 1}"},
        ])

    def test_extract_article_messages_falls_back_to_text_when_raw_missing(self):
        nodes = [
            {"role": "article", "cls": "fai-CopilotMessage bar", "text": "Copilot said:\nlegacy answer", "raw": ""},
        ]
        messages = browser_driver.extract_article_messages(nodes)
        self.assertEqual(messages, [
            {"kind": "assistant", "text": "Copilot said:\nlegacy answer"},
        ])

    def test_normalize_assistant_text_removes_placeholder_only(self):
        text = browser_driver.normalize_assistant_text("Copilot said:\nCopilot")
        self.assertEqual(text, "")

    def test_normalize_assistant_text_keeps_real_content(self):
        text = browser_driver.normalize_assistant_text("Copilot said:\nCopilot\n\n苹果")
        self.assertEqual(text, "苹果")

    def test_normalize_assistant_text_keeps_non_placeholder_first_line(self):
        text = browser_driver.normalize_assistant_text("Copilot said:\nCopilot\ncan help")
        self.assertEqual(text, "Copilot\ncan help")

    def test_normalize_user_text_flattens_whitespace_for_multiline_prompts(self):
        text = browser_driver.normalize_user_text("You said:\n请严格使用 Markdown 输出：# 标题- 一- 二```pythonprint(123)```并且不要附加解释。")
        self.assertEqual(text, "请严格使用Markdown输出：#标题-一-二```pythonprint(123)```并且不要附加解释。")

    def test_pick_response_for_prompt_matches_multiline_prompt_after_normalization(self):
        after = [
            {"kind": "user", "text": "You said:\n请严格使用 Markdown 输出：# 标题- 一- 二```pythonprint(123)```并且不要附加解释。"},
            {"kind": "assistant", "text": "Copilot said:\nCopilot\n\n# 标题\n- 一\n- 二\n```python\nprint(123)\n```"},
        ]
        prompt = "请严格使用 Markdown 输出：\n# 标题\n- 一\n- 二\n```python\nprint(123)\n```\n并且不要附加解释。"
        text = browser_driver.pick_response_for_prompt([], after, prompt)
        self.assertEqual(text, "# 标题\n- 一\n- 二\n```python\nprint(123)\n```")

    def test_pick_response_for_prompt_returns_assistant_after_matching_user(self):
        before = [
            {"kind": "user", "text": "You said:\nold"},
            {"kind": "assistant", "text": "Copilot said:\nold answer"},
        ]
        after = [
            {"kind": "user", "text": "You said:\nold"},
            {"kind": "assistant", "text": "Copilot said:\nold answer"},
            {"kind": "user", "text": "You said:\nnew prompt"},
            {"kind": "assistant", "text": "Copilot said:\nCopilot\n\nnew answer"},
        ]
        text = browser_driver.pick_response_for_prompt(before, after, "new prompt")
        self.assertEqual(text, "new answer")

    def test_pick_response_for_prompt_ignores_other_pending_reply(self):
        before = [
            {"kind": "user", "text": "You said:\nold"},
            {"kind": "assistant", "text": "Copilot said:\nold answer"},
        ]
        after = [
            {"kind": "user", "text": "You said:\nold"},
            {"kind": "assistant", "text": "Copilot said:\nold answer"},
            {"kind": "user", "text": "You said:\nother prompt"},
            {"kind": "assistant", "text": "Copilot said:\nCopilot\n\nother answer"},
            {"kind": "user", "text": "You said:\nnew prompt"},
            {"kind": "assistant", "text": "Copilot said:\nCopilot\n\nnew answer"},
        ]
        text = browser_driver.pick_response_for_prompt(before, after, "new prompt")
        self.assertEqual(text, "new answer")


class TaskpaneReadScriptTests(unittest.TestCase):
    def test_taskpane_read_script_includes_html_capture_for_markdown_content(self):
        script = browser_driver.taskpane_read_script()
        self.assertIn("innerHTML", script)

    def test_taskpane_read_script_uses_text_content_for_raw_markdown(self):
        script = browser_driver.taskpane_read_script()
        self.assertIn("announcer.textContent", script)

    def test_taskpane_read_script_reads_narrator_announcement_selector(self):
        script = browser_driver.taskpane_read_script()
        self.assertIn('[data-testid="narrator-announcement"]', script)

    def test_taskpane_read_script_scans_all_buttons_for_stop_state(self):
        script = browser_driver.taskpane_read_script()
        self.assertIn("document.querySelectorAll('button')", script)
        self.assertIn("some(btn => /stop|停止/i.test", script)

    def test_taskpane_read_script_reports_input_presence(self):
        script = browser_driver.taskpane_read_script()
        self.assertIn("hasInput", script)
        self.assertIn("!!textbox", script)

    def test_taskpane_send_script_waits_for_send_button_after_input(self):
        script = browser_driver.taskpane_send_script("hello")
        self.assertIn("new Promise", script)
        self.assertIn("setInterval", script)

    def test_taskpane_send_script_avoids_synthetic_input_and_enter_events(self):
        script = browser_driver.taskpane_send_script("hello")
        self.assertNotIn("dispatchEvent(new InputEvent", script)
        self.assertNotIn("dispatchEvent(new KeyboardEvent", script)

    def test_taskpane_reset_script_scans_multiple_button_signals(self):
        script = browser_driver.taskpane_reset_script()
        self.assertIn("innerText", script)
        self.assertIn("title", script)
        self.assertIn("aria-label", script)
        self.assertIn("data-testid", script)
        self.assertIn("document.querySelectorAll('button')", script)

    def test_taskpane_reset_script_scores_local_signals_and_excludes_close(self):
        script = browser_driver.taskpane_reset_script()
        self.assertIn("scoreSignal", script)
        self.assertIn("negativePatterns", script)
        self.assertIn("/^close$/i", script)
        self.assertIn("/^关闭$/i", script)
        self.assertIn("localTooltip", script)
        self.assertNotIn("const tooltipText = [...document.querySelectorAll", script)


class SendPromptViaTaskpaneTests(unittest.TestCase):
    def test_send_prompt_via_taskpane_awaits_async_send_script(self):
        calls = []

        class FakeClient:
            def __init__(self, ws_url):
                self.ws_url = ws_url

            def connect(self):
                pass

            def close(self):
                pass

            def call(self, method, params=None):
                calls.append((method, params))
                if method == "Runtime.enable":
                    return {}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}}
                if method == "Runtime.evaluate" and params.get("expression") == "SEND_SCRIPT":
                    return {"result": {"value": {"ok": True}}}
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_load_targets = browser_driver.load_targets
        original_pick_taskpane_target = browser_driver.pick_taskpane_target
        original_client = browser_driver.CDPClient
        original_read_script = browser_driver.taskpane_read_script
        original_send_script = browser_driver.taskpane_send_script
        try:
            browser_driver.load_targets = lambda: [{"webSocketDebuggerUrl": "ws://debug/taskpane"}]
            browser_driver.pick_taskpane_target = lambda targets: targets[0]
            browser_driver.CDPClient = FakeClient
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_send_script = lambda prompt: "SEND_SCRIPT"

            browser_driver.send_prompt_via_taskpane("hello", poll_count=1)
        finally:
            browser_driver.load_targets = original_load_targets
            browser_driver.pick_taskpane_target = original_pick_taskpane_target
            browser_driver.CDPClient = original_client
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_send_script = original_send_script

        send_call = calls[2]
        self.assertEqual(send_call[0], "Runtime.evaluate")
        self.assertTrue(send_call[1].get("awaitPromise"))

    def test_stream_prompt_via_taskpane_emits_deltas_and_returns_final_text(self):
        class FakeTime:
            def __init__(self):
                self.sleeps = []

            def sleep(self, seconds):
                self.sleeps.append(seconds)

        fake_time = FakeTime()
        read_results = iter([
            {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}},
            {"result": {"value": {"nodes": [
                {"kind": "user", "text": "hello"},
                {"kind": "assistant", "text": "正在搜索，请稍等"},
            ], "isGenerating": True, "hasInput": True}}},
            {"result": {"value": {"nodes": [
                {"kind": "user", "text": "hello"},
                {"kind": "assistant", "text": "最终答"},
            ], "isGenerating": True, "hasInput": True}}},
            {"result": {"value": {"nodes": [
                {"kind": "user", "text": "hello"},
                {"kind": "assistant", "text": "最终答案"},
            ], "isGenerating": False, "hasInput": True}}},
        ])

        class FakeClient:
            def __init__(self, ws_url):
                self.ws_url = ws_url

            def connect(self):
                pass

            def close(self):
                pass

            def call(self, method, params=None):
                if method == "Runtime.enable":
                    return {}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return next(read_results)
                if method == "Runtime.evaluate" and params.get("expression") == "SEND_SCRIPT":
                    return {"result": {"value": {"ok": True}}}
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_load_targets = browser_driver.load_targets
        original_pick_taskpane_target = browser_driver.pick_taskpane_target
        original_client = browser_driver.CDPClient
        original_read_script = browser_driver.taskpane_read_script
        original_send_script = browser_driver.taskpane_send_script
        original_extract = browser_driver.extract_article_messages
        had_time = hasattr(browser_driver, "time")
        original_time = getattr(browser_driver, "time", None)
        deltas = []
        try:
            browser_driver.load_targets = lambda: [{"webSocketDebuggerUrl": "ws://debug/taskpane"}]
            browser_driver.pick_taskpane_target = lambda targets: targets[0]
            browser_driver.CDPClient = FakeClient
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_send_script = lambda prompt: "SEND_SCRIPT"
            browser_driver.extract_article_messages = lambda nodes: nodes
            browser_driver.time = fake_time

            result = browser_driver.stream_prompt_via_taskpane("hello", poll_count=3, on_delta=deltas.append)
        finally:
            browser_driver.load_targets = original_load_targets
            browser_driver.pick_taskpane_target = original_pick_taskpane_target
            browser_driver.CDPClient = original_client
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_send_script = original_send_script
            browser_driver.extract_article_messages = original_extract
            if had_time:
                browser_driver.time = original_time
            else:
                del browser_driver.time

        self.assertEqual(result, "最终答案")
        self.assertEqual(deltas, ["最终答", "案"])
        self.assertEqual(fake_time.sleeps, [1, 1])

    def test_stream_prompt_via_taskpane_new_chat_false_keeps_existing_order(self):
        calls = []

        class FakeClient:
            def __init__(self, ws_url):
                self.ws_url = ws_url

            def connect(self):
                pass

            def close(self):
                pass

            def call(self, method, params=None):
                calls.append((method, params))
                if method == "Runtime.enable":
                    return {}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}}
                if method == "Runtime.evaluate" and params.get("expression") == "SEND_SCRIPT":
                    return {"result": {"value": {"ok": True}}}
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_load_targets = browser_driver.load_targets
        original_pick_taskpane_target = browser_driver.pick_taskpane_target
        original_client = browser_driver.CDPClient
        original_read_script = browser_driver.taskpane_read_script
        original_send_script = browser_driver.taskpane_send_script
        try:
            browser_driver.load_targets = lambda: [{"webSocketDebuggerUrl": "ws://debug/taskpane"}]
            browser_driver.pick_taskpane_target = lambda targets: targets[0]
            browser_driver.CDPClient = FakeClient
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_send_script = lambda prompt: "SEND_SCRIPT"

            browser_driver.stream_prompt_via_taskpane("hello", poll_count=1, new_chat=False)
        finally:
            browser_driver.load_targets = original_load_targets
            browser_driver.pick_taskpane_target = original_pick_taskpane_target
            browser_driver.CDPClient = original_client
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_send_script = original_send_script

        expressions = [params.get("expression") for method, params in calls if method == "Runtime.evaluate"]
        self.assertEqual(expressions, ["READ_SCRIPT", "SEND_SCRIPT", "READ_SCRIPT"])

    def test_stream_prompt_via_taskpane_new_chat_true_resets_before_send(self):
        calls = []

        class FakeClient:
            def __init__(self, ws_url):
                self.ws_url = ws_url

            def connect(self):
                pass

            def close(self):
                pass

            def call(self, method, params=None):
                calls.append((method, params))
                if method == "Runtime.enable":
                    return {}
                if method == "Runtime.evaluate" and params.get("expression") == "RESET_SCRIPT":
                    return {"result": {"value": {"ok": True, "confirmationRequired": False}}}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}}
                if method == "Runtime.evaluate" and params.get("expression") == "SEND_SCRIPT":
                    return {"result": {"value": {"ok": True}}}
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_load_targets = browser_driver.load_targets
        original_pick_taskpane_target = browser_driver.pick_taskpane_target
        original_client = browser_driver.CDPClient
        original_read_script = browser_driver.taskpane_read_script
        original_send_script = browser_driver.taskpane_send_script
        original_reset_script = browser_driver.taskpane_reset_script
        try:
            browser_driver.load_targets = lambda: [{"webSocketDebuggerUrl": "ws://debug/taskpane"}]
            browser_driver.pick_taskpane_target = lambda targets: targets[0]
            browser_driver.CDPClient = FakeClient
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_send_script = lambda prompt: "SEND_SCRIPT"
            browser_driver.taskpane_reset_script = lambda confirm=False: "CONFIRM_SCRIPT" if confirm else "RESET_SCRIPT"

            browser_driver.stream_prompt_via_taskpane("hello", poll_count=1, new_chat=True)
        finally:
            browser_driver.load_targets = original_load_targets
            browser_driver.pick_taskpane_target = original_pick_taskpane_target
            browser_driver.CDPClient = original_client
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_send_script = original_send_script
            browser_driver.taskpane_reset_script = original_reset_script

        expressions = [params.get("expression") for method, params in calls if method == "Runtime.evaluate"]
        self.assertEqual(expressions, ["RESET_SCRIPT", "READ_SCRIPT", "READ_SCRIPT", "SEND_SCRIPT", "READ_SCRIPT"])

    def test_reset_taskpane_chat_waits_until_ready_state(self):
        class FakeTime:
            def __init__(self):
                self.sleeps = []

            def sleep(self, seconds):
                self.sleeps.append(seconds)

        fake_time = FakeTime()
        calls = []
        read_results = iter([
            {"result": {"value": {"nodes": [], "isGenerating": True, "hasInput": False}}},
            {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}},
        ])

        class FakeClient:
            def call(self, method, params=None):
                calls.append((method, params))
                if method == "Runtime.evaluate" and params.get("expression") == "RESET_SCRIPT":
                    return {"result": {"value": {"ok": True, "confirmationRequired": False}}}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return next(read_results)
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_read_script = browser_driver.taskpane_read_script
        original_reset_script = browser_driver.taskpane_reset_script
        had_time = hasattr(browser_driver, "time")
        original_time = getattr(browser_driver, "time", None)
        try:
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_reset_script = lambda confirm=False: "CONFIRM_SCRIPT" if confirm else "RESET_SCRIPT"
            browser_driver.time = fake_time

            state = browser_driver.reset_taskpane_chat(FakeClient(), poll_count=3, poll_interval=0.25)
        finally:
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_reset_script = original_reset_script
            if had_time:
                browser_driver.time = original_time
            else:
                del browser_driver.time

        self.assertEqual(state, {"nodes": [], "messages": [], "is_generating": False, "has_input": True})
        self.assertEqual(fake_time.sleeps, [0.25])
        expressions = [params.get("expression") for method, params in calls if method == "Runtime.evaluate"]
        self.assertEqual(expressions, ["RESET_SCRIPT", "READ_SCRIPT", "READ_SCRIPT"])

    def test_reset_taskpane_chat_clicks_confirmation_when_prompted(self):
        calls = []

        class FakeClient:
            def call(self, method, params=None):
                calls.append((method, params))
                if method == "Runtime.evaluate" and params.get("expression") == "RESET_SCRIPT":
                    return {"result": {"value": {"ok": True, "confirmationRequired": True}}}
                if method == "Runtime.evaluate" and params.get("expression") == "CONFIRM_SCRIPT":
                    return {"result": {"value": {"ok": True}}}
                if method == "Runtime.evaluate" and params.get("expression") == "READ_SCRIPT":
                    return {"result": {"value": {"nodes": [], "isGenerating": False, "hasInput": True}}}
                raise AssertionError(f"Unexpected call: {method} {params}")

        original_read_script = browser_driver.taskpane_read_script
        original_reset_script = browser_driver.taskpane_reset_script
        try:
            browser_driver.taskpane_read_script = lambda: "READ_SCRIPT"
            browser_driver.taskpane_reset_script = lambda confirm=False: "CONFIRM_SCRIPT" if confirm else "RESET_SCRIPT"

            browser_driver.reset_taskpane_chat(FakeClient(), poll_count=1)
        finally:
            browser_driver.taskpane_read_script = original_read_script
            browser_driver.taskpane_reset_script = original_reset_script

        expressions = [params.get("expression") for method, params in calls if method == "Runtime.evaluate"]
        self.assertEqual(expressions, ["RESET_SCRIPT", "CONFIRM_SCRIPT", "READ_SCRIPT"])

    def test_send_prompt_via_taskpane_preserves_old_full_text_behavior(self):
        original_stream = browser_driver.stream_prompt_via_taskpane
        try:
            browser_driver.stream_prompt_via_taskpane = lambda prompt, poll_count=80, poll_interval=1, on_delta=None, new_chat=False: "完整回答"
            result = browser_driver.send_prompt_via_taskpane("hello")
        finally:
            browser_driver.stream_prompt_via_taskpane = original_stream

        self.assertEqual(result, "完整回答")

    def test_send_prompt_via_taskpane_passes_new_chat_flag(self):
        captured = {}
        original_stream = browser_driver.stream_prompt_via_taskpane
        try:
            def fake_stream(prompt, poll_count=80, poll_interval=1, on_delta=None, new_chat=False):
                captured.update({
                    "prompt": prompt,
                    "poll_count": poll_count,
                    "new_chat": new_chat,
                })
                return "ok"

            browser_driver.stream_prompt_via_taskpane = fake_stream
            result = browser_driver.send_prompt_via_taskpane("hello", poll_count=5, new_chat=True)
        finally:
            browser_driver.stream_prompt_via_taskpane = original_stream

        self.assertEqual(result, "ok")
        self.assertEqual(captured, {"prompt": "hello", "poll_count": 5, "new_chat": True})

    def test_reset_taskpane_chat_and_reconnect_closes_client_when_reconnect_polling_fails(self):
        reset_client = type("ResetClient", (), {"close": lambda self: None})()

        class ReconnectClient:
            def __init__(self):
                self.closed = False

            def close(self):
                self.closed = True

        reconnect_client = ReconnectClient()
        original_connect = browser_driver.connect_taskpane_client
        original_reset = browser_driver.reset_taskpane_chat
        original_load_targets = browser_driver.load_targets
        original_read = browser_driver.read_taskpane_state
        original_time = browser_driver.time
        try:
            clients = [reset_client, reconnect_client]
            browser_driver.connect_taskpane_client = lambda *args, **kwargs: clients.pop(0)
            browser_driver.reset_taskpane_chat = lambda *args, **kwargs: {"has_input": True, "is_generating": False}
            browser_driver.load_targets = lambda: [{"webSocketDebuggerUrl": "ws://debug/taskpane"}]
            browser_driver.read_taskpane_state = lambda client: (_ for _ in ()).throw(RuntimeError("poll boom"))
            browser_driver.time = type("FakeTime", (), {"sleep": staticmethod(lambda seconds: None)})()

            with self.assertRaisesRegex(RuntimeError, "poll boom"):
                browser_driver.reset_taskpane_chat_and_reconnect(poll_count=1)
        finally:
            browser_driver.connect_taskpane_client = original_connect
            browser_driver.reset_taskpane_chat = original_reset
            browser_driver.load_targets = original_load_targets
            browser_driver.read_taskpane_state = original_read
            browser_driver.time = original_time

        self.assertTrue(reconnect_client.closed)


if __name__ == "__main__":
    unittest.main()
