#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import time
import urllib.request
from pathlib import Path

import websocket

DEBUG_JSON_URL = "http://127.0.0.1:9222/json"
OUT_PATH = Path("D:/UserData/Desktop/m3652API-main/live_augloop_trace.jsonl")
WATCH_SUBSTRINGS = (
    "login.microsoftonline.com/common/oauth2/v2.0/token",
    "AcquireTokenForAugloop.ashx",
    "augloop.office.com",
)
TARGET_SUBSTRINGS = (
    "excel.cloud.microsoft/open/onedrive",
    "excel.officeapps.live.com/x/_layouts/xlviewerinternal.aspx",
    "taskpane.html",
)


def load_targets():
    with urllib.request.urlopen(DEBUG_JSON_URL, timeout=10) as resp:
        targets = json.loads(resp.read().decode("utf-8"))
    watched = []
    for target in targets:
        if target.get("type") not in ("page", "iframe"):
            continue
        url = target.get("url", "")
        if not target.get("webSocketDebuggerUrl"):
            continue
        if any(key in url for key in TARGET_SUBSTRINGS):
            watched.append(target)
    return watched


def matches_watch(url: str) -> bool:
    return any(key in (url or "") for key in WATCH_SUBSTRINGS)


def main():
    targets = load_targets()
    if not targets:
        raise SystemExit("No matching CDP targets found")

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    sockets = []
    body_request_map = {}
    next_command_id = 1000

    with OUT_PATH.open("w", encoding="utf-8") as out:
        for target in targets:
            ws = websocket.create_connection(
                target["webSocketDebuggerUrl"],
                timeout=5,
                origin="http://127.0.0.1:9222",
            )
            ws.settimeout(0.25)
            sockets.append((target, ws))
            ws.send(json.dumps({"id": 1, "method": "Page.enable", "params": {}}))
            ws.send(json.dumps({"id": 2, "method": "Network.enable", "params": {}}))

        print(f"Watching {len(sockets)} targets -> {OUT_PATH}")
        deadline = time.time() + 60

        while time.time() < deadline:
            for target, ws in sockets:
                try:
                    raw = ws.recv()
                except Exception:
                    continue

                data = json.loads(raw)
                method = data.get("method")
                params = data.get("params", {})

                if method == "Network.requestWillBeSent":
                    request = params.get("request", {})
                    url = request.get("url", "")
                    if matches_watch(url):
                        record = {
                            "kind": "request",
                            "target_url": target.get("url", ""),
                            "url": url,
                            "method": request.get("method"),
                            "headers": request.get("headers", {}),
                            "postData": request.get("postData", "")[:50000],
                            "requestId": params.get("requestId"),
                        }
                        out.write(json.dumps(record, ensure_ascii=False) + "\n")
                        out.flush()
                        print("REQ", url[:180])

                elif method == "Network.responseReceived":
                    response = params.get("response", {})
                    url = response.get("url", "")
                    request_id = params.get("requestId")
                    if matches_watch(url):
                        record = {
                            "kind": "response_meta",
                            "target_url": target.get("url", ""),
                            "url": url,
                            "status": response.get("status"),
                            "headers": response.get("headers", {}),
                            "requestId": request_id,
                        }
                        out.write(json.dumps(record, ensure_ascii=False) + "\n")
                        out.flush()
                        print("RESP", response.get("status"), url[:180])
                        next_command_id += 1
                        body_request_map[next_command_id] = {
                            "url": url,
                            "requestId": request_id,
                            "target_url": target.get("url", ""),
                        }
                        ws.send(
                            json.dumps(
                                {
                                    "id": next_command_id,
                                    "method": "Network.getResponseBody",
                                    "params": {"requestId": request_id},
                                }
                            )
                        )

                elif data.get("id") in body_request_map:
                    meta = body_request_map.pop(data["id"])
                    body = data.get("result", {}).get("body", "")
                    record = {
                        "kind": "response_body",
                        "target_url": meta["target_url"],
                        "url": meta["url"],
                        "requestId": meta["requestId"],
                        "body": body[:50000],
                    }
                    out.write(json.dumps(record, ensure_ascii=False) + "\n")
                    out.flush()
                    print("BODY", meta["url"][:120], len(body))

        print("Done")

    for _, ws in sockets:
        try:
            ws.close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
