import json
import re
import time
import urllib.request
from typing import Callable, Optional

import websocket


class CDPClient:
    def __init__(self, ws_url: str, origin: str = "http://127.0.0.1:9222"):
        self.ws_url = ws_url
        self.origin = origin
        self.ws = None
        self.next_id = 1

    def connect(self):
        self.ws = websocket.create_connection(self.ws_url, timeout=10, origin=self.origin)

    def close(self):
        if self.ws:
            self.ws.close()
            self.ws = None

    def call(self, method, params=None, timeout: float = 30):
        if not self.ws:
            raise RuntimeError("CDP client is not connected")
        msg_id = self.next_id
        self.next_id += 1
        deadline = time.monotonic() + timeout
        self.ws.send(json.dumps({"id": msg_id, "method": method, "params": params or {}}))
        while True:
            if time.monotonic() >= deadline:
                raise TimeoutError(f"CDP call timed out: {method}")
            data = json.loads(self.ws.recv())
            if data.get("id") == msg_id:
                if "error" in data:
                    raise RuntimeError(data["error"])
                return data.get("result", {})


def longest_common_prefix_length(left: str, right: str) -> int:
    limit = min(len(left), len(right))
    index = 0
    while index < limit and left[index] == right[index]:
        index += 1
    return index


def strip_placeholder_line(text: str) -> str:
    stripped = text.strip()
    if not stripped:
        return ""
    lines = stripped.splitlines()
    if not lines or lines[0].strip() != "Copilot":
        return stripped
    if len(lines) == 1:
        return ""
    if lines[1].strip() == "":
        remainder = "\n".join(lines[2:]).strip()
        return remainder
    return stripped


def remove_leading_overlap(text: str) -> str:
    for index in range(len(text)):
        candidate = text[index:]
        if candidate and text.startswith(candidate):
            return candidate
    return text


def unique_suffix_from_previous(current_text: str, previous_text: str) -> str:
    prefix_len = longest_common_prefix_length(current_text, previous_text)
    if prefix_len:
        return current_text[prefix_len:]
    for size in range(min(len(current_text), len(previous_text)), 0, -1):
        if previous_text.endswith(current_text[:size]):
            return current_text[size:]
    return current_text


def delta_tail_text(current_text: str, previous_text: str) -> str:
    delta = unique_suffix_from_previous(current_text, previous_text)
    return remove_leading_overlap(delta)


def new_text_delta(current_text: str, previous_text: str) -> str:
    if current_text.startswith(previous_text):
        return current_text[len(previous_text):]
    return delta_tail_text(current_text, previous_text)


def normalize_assistant_text(text: str) -> str:
    if not text:
        return ""
    normalized = text.strip()
    if normalized.startswith("Copilot said:\n"):
        normalized = normalized[len("Copilot said:\n"):]
    normalized = strip_placeholder_line(normalized)
    if normalized == "Copilot":
        return ""
    return normalized


def load_targets(debug_json_url: str = "http://127.0.0.1:9222/json"):
    with urllib.request.urlopen(debug_json_url, timeout=10) as resp:
        return json.loads(resp.read().decode("utf-8"))


def pick_shell_target(targets):
    candidates = [t for t in targets if t.get("type") == "page" and t.get("webSocketDebuggerUrl")]
    matches = [t for t in candidates if "excel.cloud.microsoft/open/onedrive" in t.get("url", "")]
    if not matches:
        raise RuntimeError("Excel shell page not found")
    return matches[0]


def pick_taskpane_target(targets):
    candidates = [t for t in targets if t.get("type") == "iframe" and t.get("webSocketDebuggerUrl")]
    matches = [t for t in candidates if "taskpane.html" in t.get("url", "")]
    if not matches:
        raise RuntimeError("Copilot taskpane iframe not found")
    return matches[0]


def find_interesting_globals(keys):
    allowed_prefixes = ("Office", "Excel")
    return [k for k in keys if k.startswith(allowed_prefixes)]


def pick_latest_assistant_text(messages):
    nonempty = [m for m in messages if m]
    return nonempty[-1] if nonempty else ""


def extract_article_messages(nodes):
    messages = []
    for node in nodes:
        if node.get("role") != "article":
            continue
        cls = node.get("cls", "") or ""
        text = node.get("text", "") or ""
        raw = node.get("raw", "") or ""
        if not text and not raw:
            continue
        if "fai-UserMessage" in cls:
            messages.append({"kind": "user", "text": text})
        elif "fai-CopilotMessage" in cls:
            assistant_text = raw if raw else text
            messages.append({"kind": "assistant", "text": assistant_text})
    return messages


def normalize_user_text(text: str) -> str:
    if not text:
        return ""
    text = text.strip()
    if text.startswith("You said:\n"):
        text = text[len("You said:\n"):]
    return re.sub(r"\s+", "", text)


def is_meaningful_assistant_text(text: str) -> bool:
    return bool(normalize_assistant_text(text))


def is_intermediate_assistant_text(text: str) -> bool:
    normalized = normalize_assistant_text(text)
    if not normalized:
        return False
    lowered = normalized.lower()
    patterns = [
        "正在搜索",
        "请稍等",
        "搜索中",
        "thinking",
        "searching",
        "working on it",
    ]
    return any(pattern in lowered for pattern in patterns)


def pick_new_assistant_message(before_messages, after_messages):
    if len(after_messages) <= len(before_messages):
        return ""
    appended = after_messages[len(before_messages):]
    assistants = [normalize_assistant_text(m.get("text", "")) for m in appended if m.get("kind") == "assistant"]
    assistants = [m for m in assistants if m]
    return assistants[-1] if assistants else ""


def pick_response_for_prompt(before_messages, after_messages, prompt: str):
    if len(after_messages) <= len(before_messages):
        return ""
    appended = after_messages[len(before_messages):]
    target_prompt = normalize_user_text(prompt)
    armed = False
    for message in appended:
        if message.get("kind") == "user":
            armed = normalize_user_text(message.get("text", "")) == target_prompt
            continue
        if armed and message.get("kind") == "assistant":
            text = normalize_assistant_text(message.get("text", ""))
            if text:
                return text
    return ""


def taskpane_send_script(prompt: str) -> str:
    escaped = json.dumps(prompt)
    return f"""
(() => new Promise((resolve) => {{
  const text = {escaped};
  const textbox = document.querySelector('[role="textbox"][contenteditable="true"], [role="textbox"], textarea, [contenteditable="true"], input[type="text"]');
  if (!textbox) return resolve({{ ok: false, error: 'input-not-found' }});

  textbox.focus();
  if (document.execCommand) {{
    document.execCommand('selectAll', false);
    document.execCommand('delete', false);
    document.execCommand('insertText', false, text);
  }} else {{
    textbox.textContent = text;
  }}

  let tries = 0;
  const timer = setInterval(() => {{
    tries += 1;
    const sendButton = document.querySelector('.fai-SendButton') || [...document.querySelectorAll('button')].find(btn => /send|发送/i.test((btn.title || btn.getAttribute('aria-label') || btn.innerText || '').trim()));
    if (sendButton) {{
      clearInterval(timer);
      sendButton.click();
      return resolve({{ ok: true, mode: 'button', tries }});
    }}
    if (tries >= 20) {{
      clearInterval(timer);
      resolve({{ ok: false, error: 'send-button-not-found', tries }});
    }}
  }}, 100);
}}))()
"""


def taskpane_reset_script(confirm: bool = False) -> str:
    mode = json.dumps("confirm" if confirm else "start")
    return f"""
(() => new Promise((resolve) => {{
  const mode = {mode};
  const actionPatterns = mode === 'confirm'
    ? [/^确定$/i, /^确认$/i, /^继续$/i, /^yes$/i, /confirm/i, /start\\s+new\\s+chat/i, /new\\s+chat/i, /开始新的聊天/i, /开始新聊天/i]
    : [/new\\s+chat/i, /new\\s+conversation/i, /start\\s+new\\s+chat/i, /restart/i, /开始新的聊天/i, /开始新聊天/i, /新建聊天/i, /新聊天/i, /重新开始/i, /重新对话/i];
  const negativePatterns = mode === 'confirm'
    ? [/^close$/i, /^关闭$/i]
    : [/^close$/i, /^关闭$/i, /dismiss/i, /取消/i];
  const normalize = (value) => (value || '').toString().trim();
  const collectSignals = (button) => {{
    const title = normalize(button.getAttribute('title') || button.title);
    const ariaLabel = normalize(button.getAttribute('aria-label'));
    const innerText = normalize(button.innerText || button.textContent);
    const dataTestId = normalize(button.getAttribute('data-testid'));
    const className = normalize(button.className);
    const nearbyText = normalize(
      (button.closest('[role="dialog"], [role="menu"], [role="toolbar"], [data-testid], .fai-ChatHeader, .fai-OverflowButton, .ms-Callout') || button.parentElement || {{}}).innerText
    );
    const localTooltip = normalize(
      (button.closest('[role="toolbar"], [data-testid], .fai-ChatHeader, .ms-Callout') || button.parentElement || {{}}).querySelector?.('[role="tooltip"], .ms-Tooltip, [data-tippy-root]')?.innerText
    );
    return {{ innerText, title, ariaLabel, dataTestId, className, nearbyText, localTooltip }};
  }};
  const scoreSignal = (key) => {{
    if (key === 'ariaLabel') return 8;
    if (key === 'title') return 7;
    if (key === 'innerText') return 6;
    if (key === 'dataTestId') return 5;
    if (key === 'nearbyText') return 2;
    if (key === 'localTooltip') return 1;
    if (key === 'className') return 1;
    return 0;
  }};
  const matchSignals = (signals, patterns) => {{
    const entries = Object.entries(signals).filter(([, value]) => value);
    return entries
      .filter(([, value]) => patterns.some(pattern => pattern.test(value)))
      .map(([key]) => key);
  }};
  const hasNegativeSignal = (signals) => Object.values(signals)
    .filter(Boolean)
    .some(value => negativePatterns.some(pattern => pattern.test(value)));
  const buttons = [...document.querySelectorAll('button')].map((button, index) => {{
    const signals = collectSignals(button);
    const matchedSignals = matchSignals(signals, actionPatterns);
    const excluded = hasNegativeSignal(signals);
    const score = matchedSignals.reduce((total, key) => total + scoreSignal(key), 0);
    return {{
      button,
      index,
      signals,
      matchedSignals,
      excluded,
      score,
      summary: {{
        index,
        innerText: signals.innerText,
        title: signals.title,
        ariaLabel: signals.ariaLabel,
        dataTestId: signals.dataTestId,
        className: signals.className,
        nearbyText: signals.nearbyText,
        localTooltip: signals.localTooltip,
      }},
    }};
  }});
  const candidates = buttons
    .filter(info => info.matchedSignals.length > 0 && !info.excluded)
    .sort((left, right) => right.score - left.score || right.matchedSignals.length - left.matchedSignals.length || left.index - right.index);
  if (!candidates.length) {{
    return resolve({{
      ok: false,
      hit: false,
      mode,
      clues: ['button', 'innerText', 'title', 'aria-label', 'data-testid', 'class'],
      clicked: null,
      confirmationButton: null,
      candidates: buttons.map(info => info.summary),
    }});
  }}

  const chosen = candidates[0];
  chosen.button.click();

  setTimeout(() => {{
    const confirmPatterns = [/^确定$/i, /^确认$/i, /^继续$/i, /^yes$/i, /confirm/i, /start\\s+new\\s+chat/i, /new\\s+chat/i, /开始新的聊天/i, /开始新聊天/i];
    const confirmCandidates = [...document.querySelectorAll('button')]
      .map((button, index) => {{
        const signals = collectSignals(button);
        const matchedSignals = matchSignals(signals, confirmPatterns);
        const excluded = hasNegativeSignal(signals);
        const score = matchedSignals.reduce((total, key) => total + scoreSignal(key), 0);
        return {{ index, matchedSignals, excluded, score, summary: {{
          index,
          innerText: signals.innerText,
          title: signals.title,
          ariaLabel: signals.ariaLabel,
          dataTestId: signals.dataTestId,
          className: signals.className,
          nearbyText: signals.nearbyText,
          localTooltip: signals.localTooltip,
        }} }};
      }})
      .filter(info => info.matchedSignals.length > 0 && !info.excluded)
      .sort((left, right) => right.score - left.score || right.matchedSignals.length - left.matchedSignals.length || left.index - right.index);
    resolve({{
      ok: true,
      hit: true,
      mode,
      clues: chosen.matchedSignals,
      clicked: chosen.summary,
      confirmationButton: confirmCandidates.length ? confirmCandidates[0].summary : null,
      confirmationRequired: mode !== 'confirm' && confirmCandidates.length > 0,
    }});
  }}, 150);
}}))()
"""


def taskpane_read_script() -> str:
    return """
(() => {
  const textbox = document.querySelector('[role="textbox"][contenteditable="true"], [role="textbox"], textarea, [contenteditable="true"], input[type="text"]');
  const nodes = [...document.querySelectorAll('[role="article"], [role="textbox"], .fai-CopilotMessage, .fai-UserMessage')];
  const buttons = [...document.querySelectorAll('button')];
  const isGenerating = buttons.some(btn => /stop|停止/i.test(((btn.getAttribute('title') || btn.getAttribute('aria-label') || btn.innerText || '').trim())));
  return {
    isGenerating,
    hasInput: !!textbox,
    nodes: nodes.map((node) => ({
      role: node.getAttribute('role'),
      cls: (node.className || '').toString(),
      text: (node.innerText || node.textContent || '').trim(),
      raw: (() => {
        const announcer = node.querySelector && node.querySelector('[data-testid="narrator-announcement"]');
        return announcer ? (announcer.textContent || announcer.innerText || '').trim() : '';
      })(),
      html: (node.innerHTML || '').trim()
    }))
  };
})()
"""


def ensure_browser_ready():
    targets = load_targets()
    pick_shell_target(targets)
    pick_taskpane_target(targets)
    return targets


def read_taskpane_state(client):
    result = client.call("Runtime.evaluate", {"expression": taskpane_read_script(), "returnByValue": True})
    value = result.get("result", {}).get("value", {})
    nodes = value.get("nodes", [])
    return {
        "nodes": nodes,
        "messages": extract_article_messages(nodes),
        "is_generating": value.get("isGenerating", False),
        "has_input": value.get("hasInput", False),
    }


def connect_taskpane_client(targets=None, retries: int = 2, retry_delay: float = 1.0):
    last_error = None
    current_targets = targets
    for attempt in range(retries + 1):
        client = None
        try:
            if current_targets is None:
                current_targets = load_targets()
            taskpane = pick_taskpane_target(current_targets)
            client = CDPClient(taskpane["webSocketDebuggerUrl"])
            client.connect()
            client.call("Runtime.enable")
            return client
        except Exception as exc:
            last_error = exc
            if client:
                try:
                    client.close()
                except Exception:
                    pass
            if attempt >= retries:
                break
            time.sleep(retry_delay)
            current_targets = load_targets()
    raise last_error


def reset_taskpane_chat(client, poll_count: int = 20, poll_interval: float = 0.5):
    result = client.call(
        "Runtime.evaluate",
        {"expression": taskpane_reset_script(), "returnByValue": True, "awaitPromise": True},
    )
    value = result.get("result", {}).get("value", {})
    if not value.get("ok"):
        raise RuntimeError(value.get("error", "taskpane-reset-failed"))

    if value.get("confirmationRequired"):
        confirm_result = client.call(
            "Runtime.evaluate",
            {"expression": taskpane_reset_script(confirm=True), "returnByValue": True, "awaitPromise": True},
        )
        confirm_value = confirm_result.get("result", {}).get("value", {})
        if not confirm_value.get("ok"):
            raise RuntimeError(confirm_value.get("error", "taskpane-reset-confirmation-failed"))

    stable_state = None
    for attempt in range(poll_count):
        if attempt:
            time.sleep(poll_interval)
        state = read_taskpane_state(client)
        stable_state = state
        if state["has_input"] and not state["is_generating"]:
            return state

    raise RuntimeError(f"taskpane-reset-did-not-stabilize: {stable_state}")


def reset_taskpane_chat_and_reconnect(targets=None, poll_count: int = 20, poll_interval: float = 0.5, reconnect_retries: int = 3, reconnect_delay: float = 1.0):
    reset_client = connect_taskpane_client(targets=targets)
    try:
        reset_taskpane_chat(reset_client, poll_count=poll_count, poll_interval=poll_interval)
    finally:
        reset_client.close()

    time.sleep(reconnect_delay)
    refreshed_targets = load_targets()
    client = connect_taskpane_client(targets=refreshed_targets, retries=reconnect_retries, retry_delay=reconnect_delay)
    state = None
    try:
        for attempt in range(poll_count):
            if attempt:
                time.sleep(poll_interval)
            state = read_taskpane_state(client)
            if state["has_input"] and not state["is_generating"]:
                return client, state
    except Exception:
        client.close()
        raise

    client.close()
    raise RuntimeError(f"taskpane-reset-reconnect-did-not-stabilize: {state}")


def stream_prompt_via_taskpane(
    prompt: str,
    poll_count: int = 80,
    poll_interval: float = 1,
    on_delta: Optional[Callable[[str], None]] = None,
    new_chat: bool = False,
):
    targets = load_targets()
    if new_chat:
        client, before_state = reset_taskpane_chat_and_reconnect(targets=targets)
    else:
        client = connect_taskpane_client(targets=targets)
        before_state = read_taskpane_state(client)
    try:
        before_messages = before_state["messages"]

        client.call("Runtime.evaluate", {"expression": taskpane_send_script(prompt), "returnByValue": True, "awaitPromise": True})

        latest_assistant = ""
        for attempt in range(poll_count):
            if attempt:
                time.sleep(poll_interval)
            state = read_taskpane_state(client)
            candidate = pick_response_for_prompt(before_messages, state["messages"], prompt)
            if candidate and candidate != latest_assistant and not is_intermediate_assistant_text(candidate):
                delta = new_text_delta(candidate, latest_assistant)
                latest_assistant = candidate
                if delta and on_delta:
                    on_delta(delta)
            if candidate and is_meaningful_assistant_text(candidate) and not state["is_generating"]:
                break
        return latest_assistant
    finally:
        client.close()


def send_prompt_via_taskpane(prompt: str, poll_count: int = 80, new_chat: bool = False):
    return stream_prompt_via_taskpane(prompt, poll_count=poll_count, new_chat=new_chat)
