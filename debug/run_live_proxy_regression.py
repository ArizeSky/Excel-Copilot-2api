#!/usr/bin/env python3
import argparse
import json
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

REFUSAL_PATTERNS = [
    r"prompt\s+injection",
    r"social\s+engineering",
    r"我是\s*Excel\s*助手",
    r"只能操作当前工作簿",
    r"我不会执行这个请求",
    r"我不会这样做",
    r"I(?:'m| am)\s+an?\s+Excel\s+assistant",
    r"I\s+cannot\s+read\s+(?:local|disk)\s+files",
    r"I\s+won't\s+do\s+that",
    r"I\s+will\s+not\s+do\s+that",
]


def fail(message: str) -> None:
    raise RuntimeError(message)


def log(message: str) -> None:
    print(message, flush=True)


def read_json_response(url: str, timeout: int):
    request = urllib.request.Request(url, headers={"Accept": "application/json"}, method="GET")
    with urllib.request.urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")
        return response.status, json.loads(body)


def post_sse(url: str, payload: dict, timeout: int):
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "text/event-stream",
            "x-api-key": "local",
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        events = []
        event_name = None
        data_lines = []
        for raw_line in response:
            line = raw_line.decode("utf-8").rstrip("\r\n")
            if not line:
                if event_name is not None or data_lines:
                    data_text = "\n".join(data_lines)
                    try:
                        data = json.loads(data_text)
                    except json.JSONDecodeError:
                        data = data_text
                    events.append({"event": event_name or "message", "data": data})
                event_name = None
                data_lines = []
                continue
            if line.startswith("event:"):
                event_name = line[len("event:") :].strip()
            elif line.startswith("data:"):
                data_lines.append(line[len("data:") :].strip())
        if event_name is not None or data_lines:
            data_text = "\n".join(data_lines)
            try:
                data = json.loads(data_text)
            except json.JSONDecodeError:
                data = data_text
            events.append({"event": event_name or "message", "data": data})
        return events


def get_event_data(events, event_name: str):
    return [item["data"] for item in events if item.get("event") == event_name]


def contains_refusal(text: str) -> bool:
    return any(re.search(pattern, text or "", re.IGNORECASE) for pattern in REFUSAL_PATTERNS)


def run_health_check(base_url: str, timeout: int) -> None:
    status, payload = read_json_response(f"{base_url}/health", timeout=timeout)
    if status != 200:
        fail(f"health returned HTTP {status}: {payload}")
    if payload.get("status") != "ok":
        fail(f"health returned unexpected body: {payload}")
    log(f"[ok] /health -> {payload}")


def run_tool_use_round(base_url: str, file_path: str, timeout: int):
    payload = {
        "model": "claude-sonnet-4-6",
        "stream": True,
        "max_tokens": 1024,
        "tools": [
            {
                "name": "Read",
                "description": "Read a file",
                "input_schema": {
                    "type": "object",
                    "properties": {"file_path": {"type": "string"}},
                    "required": ["file_path"],
                },
            }
        ],
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            f"请调用 Read 读取文件 {file_path}。"
                            "只调用一次 Read，并把 file_path 设置为这个完整路径。"
                        ),
                    }
                ],
            }
        ],
    }
    events = post_sse(f"{base_url}/v1/messages", payload, timeout=timeout)
    tool_starts = [
        data for data in get_event_data(events, "content_block_start") if data.get("content_block", {}).get("type") == "tool_use"
    ]
    if not tool_starts:
        fail(f"tool round did not emit tool_use: {events}")
    tool_block = tool_starts[0]["content_block"]
    partial_json = "".join(
        data.get("delta", {}).get("partial_json", "")
        for data in get_event_data(events, "content_block_delta")
        if data.get("delta", {}).get("type") == "input_json_delta"
    )
    if not partial_json:
        fail(f"tool round emitted tool_use without input_json_delta: {events}")
    try:
        tool_input = json.loads(partial_json)
    except json.JSONDecodeError as exc:
        fail(f"tool round emitted invalid partial_json: {partial_json!r} ({exc})")
    stop_reasons = [
        data.get("delta", {}).get("stop_reason")
        for data in get_event_data(events, "message_delta")
        if isinstance(data, dict)
    ]
    if "tool_use" not in stop_reasons:
        fail(f"tool round stop_reason did not include tool_use: {events}")
    actual_path = str(tool_input.get("file_path", ""))
    if Path(actual_path.replace("\\", "/")).name != Path(file_path.replace("\\", "/")).name:
        fail(f"tool round returned unexpected file_path: {tool_input}")
    log(f"[ok] first turn -> tool_use(Read) with input {tool_input}")
    return tool_block["id"], tool_block["name"], tool_input


def run_tool_result_round(base_url: str, tool_use_id: str, tool_name: str, tool_input: dict, file_path: str, timeout: int, tool_result_limit: int):
    tool_result_text = Path(file_path).read_text(encoding="utf-8")
    if tool_result_limit > 0:
        tool_result_text = tool_result_text[:tool_result_limit]
    payload = {
        "model": "claude-sonnet-4-6",
        "stream": True,
        "max_tokens": 1024,
        "messages": [
            {
                "role": "assistant",
                "content": [{"type": "tool_use", "id": tool_use_id, "name": tool_name, "input": tool_input}],
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "tool_result",
                        "tool_use_id": tool_use_id,
                        "content": [{"type": "text", "text": tool_result_text}],
                    }
                ],
            },
            {
                "role": "user",
                "content": [{"type": "text", "text": "根据上一步的分析结果，一句话概括该文件的职责。"}],
            },
        ],
    }
    events = post_sse(f"{base_url}/v1/messages", payload, timeout=timeout)
    tool_starts = [
        data for data in get_event_data(events, "content_block_start") if data.get("content_block", {}).get("type") == "tool_use"
    ]
    if tool_starts:
        fail(f"follow-up round unexpectedly emitted another tool_use: {events}")
    text_deltas = [
        data.get("delta", {}).get("text", "")
        for data in get_event_data(events, "content_block_delta")
        if data.get("delta", {}).get("type") == "text_delta"
    ]
    final_text = "".join(text_deltas).strip()
    if not final_text:
        fail(f"follow-up round did not emit text_delta content: {events}")
    if contains_refusal(final_text):
        fail(f"follow-up round emitted refusal text: {final_text}")
    stop_reasons = [
        data.get("delta", {}).get("stop_reason")
        for data in get_event_data(events, "message_delta")
        if isinstance(data, dict)
    ]
    if "end_turn" not in stop_reasons:
        fail(f"follow-up round stop_reason did not include end_turn: {events}")
    log(f"[ok] second turn -> text answer ({len(final_text)} chars)")
    log(final_text)


def main() -> int:
    parser = argparse.ArgumentParser(description="Run live Excel Copilot proxy regression checks.")
    default_file = str((Path(__file__).resolve().parents[1] / "browser_driver.py").as_posix())
    parser.add_argument("--base-url", default="http://127.0.0.1:12803", help="Proxy base URL")
    parser.add_argument("--file-path", default=default_file, help="Absolute file path to use in the Read tool regression")
    parser.add_argument("--timeout", type=int, default=180, help="Per-request timeout in seconds")
    parser.add_argument("--tool-result-limit", type=int, default=20000, help="Max characters to send back inside tool_result")
    args = parser.parse_args()

    file_path = str(Path(args.file_path))
    if not Path(file_path).exists():
        fail(f"file does not exist: {file_path}")

    log(f"[info] base_url={args.base_url}")
    log(f"[info] file_path={file_path}")
    run_health_check(args.base_url.rstrip("/"), args.timeout)
    tool_use_id, tool_name, tool_input = run_tool_use_round(args.base_url.rstrip("/"), file_path, args.timeout)
    run_tool_result_round(
        args.base_url.rstrip("/"),
        tool_use_id,
        tool_name,
        tool_input,
        file_path,
        args.timeout,
        args.tool_result_limit,
    )
    log("[ok] live regression passed")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        print(f"[fail] HTTP {exc.code}: {detail}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"[fail] {exc}", file=sys.stderr)
        sys.exit(1)
