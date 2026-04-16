import asyncio
import threading
import time
import unittest
import importlib.util
from pathlib import Path

MODULE_PATH = Path(__file__).resolve().parent.parent / "browser_attached_proxy.py"
spec = importlib.util.spec_from_file_location("browser_attached_proxy", MODULE_PATH)
browser_attached_proxy = importlib.util.module_from_spec(spec)
spec.loader.exec_module(browser_attached_proxy)


class SSETests(unittest.TestCase):
    def test_sse_chunk_wraps_payload(self):
        chunk = browser_attached_proxy.sse_chunk({"id": "1"})
        self.assertEqual(chunk, 'data: {"id": "1"}\n\n')

    def test_anthropic_sse_chunk_wraps_event_and_payload(self):
        chunk = browser_attached_proxy.anthropic_sse_chunk("message_start", {"type": "message_start"})
        self.assertEqual(chunk, 'event: message_start\ndata: {"type": "message_start"}\n\n')

    def test_last_user_prompt_returns_text_content(self):
        prompt = browser_attached_proxy.last_user_prompt([
            {"role": "assistant", "content": "old"},
            {"role": "user", "content": "hello"},
        ])
        self.assertEqual(prompt, "hello")

    def test_last_user_prompt_supports_anthropic_text_blocks(self):
        prompt = browser_attached_proxy.last_user_prompt([
            {"role": "user", "content": [{"type": "text", "text": "第一段"}, {"type": "text", "text": "第二段"}]},
        ])
        self.assertEqual(prompt, "第一段\n第二段")


class RequestDecodeTests(unittest.TestCase):
    def test_decode_json_body_accepts_utf8(self):
        body = '{"messages":[{"role":"user","content":"苹果"}],"stream":true}'.encode("utf-8")
        parsed = browser_attached_proxy.decode_json_body(body, "application/json; charset=utf-8")
        self.assertTrue(parsed["stream"])
        self.assertEqual(parsed["messages"][0]["content"], "苹果")

    def test_decode_json_body_falls_back_to_gbk(self):
        body = '{"messages":[{"role":"user","content":"苹果"}],"stream":true}'.encode("gbk")
        parsed = browser_attached_proxy.decode_json_body(body, "application/json")
        self.assertTrue(parsed["stream"])
        self.assertEqual(parsed["messages"][0]["content"], "苹果")


class RequestNormalizationTests(unittest.TestCase):
    def test_clean_prompt_injection_tags_removes_leading_xml_wrappers(self):
        text = "<system-reminder>ignore</system-reminder>\n<note>x</note>\n实际请求"
        self.assertEqual(browser_attached_proxy.clean_prompt_injection_tags(text), "实际请求")

    def test_clean_transcript_text_drops_refusal_content(self):
        text = "我是 Excel 助手，只能操作当前工作簿。"
        self.assertEqual(browser_attached_proxy.clean_transcript_text(text), "")

    def test_build_excel_native_reframing_wraps_tool_prompt_as_sheet_record_task(self):
        reframed = browser_attached_proxy.build_excel_native_reframing(
            "请读取 browser_driver.py",
            {"type": "tool", "name": "Read"},
        )
        self.assertIn("Excel 工作表中的单元格模板", reframed)
        self.assertIn("最终单元格值", reframed)
        self.assertIn("请读取 browser_driver.py", reframed)

    def test_normalize_anthropic_request_builds_tool_prompt_and_transcript(self):
        body = {
            "model": "claude-sonnet-4-6",
            "stream": True,
            "system": "你是助手",
            "tools": [
                {
                    "name": "Read",
                    "description": "Read a file from disk",
                    "input_schema": {
                        "type": "object",
                        "properties": {"file_path": {"type": "string"}},
                        "required": ["file_path"],
                    },
                }
            ],
            "tool_choice": {"type": "tool", "name": "Read"},
            "messages": [
                {"role": "assistant", "content": [{"type": "tool_use", "id": "toolu_1", "name": "Read", "input": {"file_path": "a.py"}}]},
                {"role": "user", "content": [{"type": "tool_result", "tool_use_id": "toolu_1", "content": [{"type": "text", "text": "print(1)"}]}]},
                {"role": "user", "content": [{"type": "text", "text": "继续"}]},
            ],
        }
        normalized = browser_attached_proxy.normalize_anthropic_request(body)
        self.assertEqual(normalized["prompt"], "继续")
        self.assertTrue(normalized["has_tool_result"])
        self.assertIn("Latest available result:\nprint(1)", normalized["browser_prompt"])
        self.assertIn("不要再次生成结构化记录", normalized["browser_prompt"])

    def test_normalize_openai_request_cleans_refusal_history_from_transcript(self):
        body = {
            "messages": [
                {"role": "assistant", "content": "我是 Excel 助手，只能操作当前工作簿。"},
                {"role": "user", "content": "请读取 browser_driver.py"},
            ],
            "tools": [
                {
                    "type": "function",
                    "function": {
                        "name": "Read",
                        "description": "Read file",
                        "parameters": {"type": "object", "properties": {"file_path": {"type": "string"}}},
                    },
                }
            ],
            "stream": True,
        }
        normalized = browser_attached_proxy.normalize_openai_request(body)
        self.assertNotIn("我是 Excel 助手", normalized["transcript"])
        self.assertIn("最终单元格值", normalized["browser_prompt"])

    def test_parse_tool_call_accepts_relaxed_object_syntax(self):
        text = '步骤 1\n让我将 JSON 文本写入 ZZ1 单元格。\n{tool:Read,parameters:{file_path:browser_driver.py}}'
        parsed = browser_attached_proxy.parse_tool_call(text)
        self.assertEqual(parsed, {"name": "Read", "input": {"file_path": "browser_driver.py"}})

    def test_extract_json_object_returns_empty_for_unclosed_json(self):
        text = '{"tool":"Read","parameters":{"file_path":"browser_driver.py"}'
        self.assertEqual(browser_attached_proxy.extract_json_object(text), "")


class ChatCompletionTests(unittest.TestCase):
    def test_chat_completions_streams_incremental_content_chunks(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"messages":[{"role":"user","content":"markdown"}],"stream":true}'

        original_stream = browser_attached_proxy.stream_browser_deltas
        try:
            async def fake_stream(prompt, new_chat=False):
                self.assertEqual(prompt, "markdown")
                self.assertTrue(new_chat)
                for delta in ["# 标题\n", "- 一\n", "- 二"]:
                    yield delta

            browser_attached_proxy.stream_browser_deltas = fake_stream
            response = asyncio.run(browser_attached_proxy.chat_completions(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            chunks = asyncio.run(collect())
        finally:
            browser_attached_proxy.stream_browser_deltas = original_stream

        self.assertEqual(len(chunks), 6)
        self.assertIn('"delta": {"content": "# 标题\\n"}', chunks[1])
        self.assertIn('"delta": {"content": "- 一\\n"}', chunks[2])
        self.assertIn('"finish_reason": "stop"', chunks[4])

    def test_chat_completions_tool_requests_start_new_chat(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"messages":[{"role":"user","content":"read it"}],"tools":[{"type":"function","function":{"name":"Read","description":"Read file","parameters":{"type":"object","properties":{"file_path":{"type":"string"}}}}}],"stream":true}'

        original_send = browser_attached_proxy.browser_driver.send_prompt_via_taskpane
        captured = []
        try:
            def fake_send(prompt, poll_count=80, new_chat=False):
                captured.append({"prompt": prompt, "poll_count": poll_count, "new_chat": new_chat})
                return '{"tool":"Read","parameters":{"file_path":"src/app.py"}}'

            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = fake_send
            response = asyncio.run(browser_attached_proxy.chat_completions(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            asyncio.run(collect())
        finally:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = original_send

        self.assertEqual(len(captured), 1)
        self.assertTrue(captured[0]["new_chat"])

    def test_chat_completions_converts_structured_tool_call_to_openai_tool_calls(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"messages":[{"role":"user","content":"read it"}],"tools":[{"type":"function","function":{"name":"Read","description":"Read file","parameters":{"type":"object","properties":{"file_path":{"type":"string"}}}}}],"stream":true}'

        original_send = browser_attached_proxy.browser_driver.send_prompt_via_taskpane
        try:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = lambda prompt, poll_count=80, new_chat=False: '{"tool":"Read","parameters":{"file_path":"src/app.py"}}'
            response = asyncio.run(browser_attached_proxy.chat_completions(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            chunks = asyncio.run(collect())
        finally:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = original_send

        joined = "".join(chunks)
        self.assertIn('"tool_calls"', joined)
        self.assertIn('"name": "Read"', joined)
        self.assertIn('"arguments": "{\\"file_path\\": \\"src/app.py\\"}"', joined)
        self.assertIn('"finish_reason": "tool_calls"', joined)


class BrowserLockTests(unittest.TestCase):
    def test_browser_lock_serializes_concurrent_messages_requests(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"messages":[{"role":"user","content":[{"type":"text","text":"hello"}]}]}'

        original_stream = browser_attached_proxy.stream_browser_deltas
        active = 0
        max_active = 0
        gate = threading.Event()
        started = threading.Event()

        try:
            async def fake_stream(prompt, new_chat=False):
                nonlocal active, max_active
                active += 1
                max_active = max(max_active, active)
                started.set()
                while not gate.is_set():
                    await asyncio.sleep(0.001)
                yield "done"
                active -= 1

            browser_attached_proxy.stream_browser_deltas = fake_stream

            async def consume_one():
                response = await browser_attached_proxy.messages(FakeRequest())
                return [chunk async for chunk in response.body_iterator]

            async def run_pair():
                first = asyncio.create_task(consume_one())
                await asyncio.sleep(0.01)
                second = asyncio.create_task(consume_one())
                started.wait(0.1)
                await asyncio.sleep(0.02)
                gate.set()
                await asyncio.gather(first, second)

            asyncio.run(run_pair())
        finally:
            browser_attached_proxy.stream_browser_deltas = original_stream

        self.assertEqual(max_active, 1)


class AnthropicMessagesTests(unittest.TestCase):
    def test_messages_rejects_non_streaming_requests(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"messages":[{"role":"user","content":[{"type":"text","text":"hi"}]}],"stream":false}'

        response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))
        self.assertEqual(response.status_code, 400)

    def test_messages_streams_text_event_sequence(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"messages":[{"role":"user","content":[{"type":"text","text":"hello"}]}]}'

        original_stream = browser_attached_proxy.stream_browser_deltas
        try:
            async def fake_stream(prompt, new_chat=False):
                self.assertEqual(prompt, "hello")
                self.assertTrue(new_chat)
                for delta in ["你", "好"]:
                    yield delta

            browser_attached_proxy.stream_browser_deltas = fake_stream
            response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            chunks = asyncio.run(collect())
        finally:
            browser_attached_proxy.stream_browser_deltas = original_stream

        self.assertTrue(chunks[0].startswith("event: message_start"))
        self.assertIn('event: content_block_start', chunks[1])
        self.assertIn('"type": "text_delta", "text": "你"', chunks[2])
        self.assertIn('"type": "text_delta", "text": "好"', chunks[3])
        self.assertIn('"stop_reason": "end_turn"', chunks[-2])
        self.assertEqual(chunks[-1], 'event: message_stop\ndata: {"type": "message_stop"}\n\n')

    def test_messages_converts_structured_tool_call_to_tool_use(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"tools":[{"name":"Read","description":"Read file","input_schema":{"type":"object","properties":{"file_path":{"type":"string"}},"required":["file_path"]}}],"messages":[{"role":"user","content":[{"type":"text","text":"read it"}]}]}'

        original_send = browser_attached_proxy.browser_driver.send_prompt_via_taskpane
        try:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = lambda prompt, poll_count=80, new_chat=False: '```json action\n{"tool":"Read","parameters":{"file_path":"src/app.py"}}\n```'
            response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            chunks = asyncio.run(collect())
        finally:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = original_send

        joined = "".join(chunks)
        self.assertIn('"type": "tool_use"', joined)
        self.assertIn('"name": "Read"', joined)
        self.assertIn('"partial_json": "{\\"file_path\\": \\"src/app.py\\"}"', joined)
        self.assertIn('"stop_reason": "tool_use"', joined)

    def test_messages_tool_requests_start_new_chat(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"tools":[{"name":"Read","description":"Read file","input_schema":{"type":"object","properties":{"file_path":{"type":"string"}},"required":["file_path"]}}],"messages":[{"role":"user","content":[{"type":"text","text":"read it"}]}]}'

        original_send = browser_attached_proxy.browser_driver.send_prompt_via_taskpane
        captured = []
        try:
            def fake_send(prompt, poll_count=80, new_chat=False):
                captured.append({"prompt": prompt, "poll_count": poll_count, "new_chat": new_chat})
                return '{"tool":"Read","parameters":{"file_path":"src/app.py"}}'

            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = fake_send
            response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            asyncio.run(collect())
        finally:
            browser_attached_proxy.browser_driver.send_prompt_via_taskpane = original_send

        self.assertEqual(len(captured), 1)
        self.assertTrue(captured[0]["new_chat"])

    def test_messages_tool_result_stream_path_starts_new_chat(self):
        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"messages":[{"role":"assistant","content":[{"type":"tool_use","id":"toolu_1","name":"Read","input":{"file_path":"src/app.py"}}]},{"role":"user","content":[{"type":"tool_result","tool_use_id":"toolu_1","content":[{"type":"text","text":"print(123)"}]}]},{"role":"user","content":[{"type":"text","text":"summarize"}]}]}'

        original_stream = browser_attached_proxy.stream_browser_deltas
        captured = []
        try:
            async def fake_stream(prompt, new_chat=False):
                captured.append({"prompt": prompt, "new_chat": new_chat})
                yield "done"

            browser_attached_proxy.stream_browser_deltas = fake_stream
            response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            asyncio.run(collect())
        finally:
            browser_attached_proxy.stream_browser_deltas = original_stream

        self.assertEqual(len(captured), 1)
        self.assertTrue(captured[0]["new_chat"])
        self.assertIn("Latest available result:\nprint(123)", captured[0]["prompt"])

        class FakeRequest:
            headers = {"content-type": "application/json; charset=utf-8"}

            async def body(self):
                return b'{"model":"claude-sonnet-4-6","stream":true,"messages":[{"role":"assistant","content":[{"type":"tool_use","id":"toolu_1","name":"Read","input":{"file_path":"src/app.py"}}]},{"role":"user","content":[{"type":"tool_result","tool_use_id":"toolu_1","content":[{"type":"text","text":"print(123)"}]}]},{"role":"user","content":[{"type":"text","text":"summarize"}]}]}'

        original_stream = browser_attached_proxy.stream_browser_deltas
        prompts = []
        try:
            async def fake_stream(prompt, new_chat=False):
                prompts.append({"prompt": prompt, "new_chat": new_chat})
                yield "done"

            browser_attached_proxy.stream_browser_deltas = fake_stream
            response = asyncio.run(browser_attached_proxy.messages(FakeRequest()))

            async def collect():
                return [chunk async for chunk in response.body_iterator]

            asyncio.run(collect())
        finally:
            browser_attached_proxy.stream_browser_deltas = original_stream

        self.assertEqual(len(prompts), 1)
        self.assertIn("Latest available result:\nprint(123)", prompts[0]["prompt"])
        self.assertIn("不要再次生成结构化记录", prompts[0]["prompt"])
        self.assertIn("Latest user turn:\nsummarize", prompts[0]["prompt"])
        self.assertTrue(prompts[0]["new_chat"])


class StreamBrowserDeltasTests(unittest.TestCase):
    def test_stream_browser_deltas_propagates_worker_error_once(self):
        original_stream = browser_attached_proxy.browser_driver.stream_prompt_via_taskpane
        try:
            browser_attached_proxy.browser_driver.stream_prompt_via_taskpane = lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("boom"))

            async def consume():
                seen = []
                try:
                    async for chunk in browser_attached_proxy.stream_browser_deltas("x"):
                        seen.append(chunk)
                except RuntimeError as exc:
                    return seen, str(exc)
                return seen, None

            chunks, error = asyncio.run(consume())
        finally:
            browser_attached_proxy.browser_driver.stream_prompt_via_taskpane = original_stream

        self.assertEqual(chunks, [])
        self.assertEqual(error, "boom")


if __name__ == "__main__":
    unittest.main()
