#!/usr/bin/env python3
"""Crowdin 文件翻译同步到 APITable 留言列的本地网页工具。"""

from __future__ import annotations

import io
import json
import mimetypes
import os
import re
import socket
import ssl
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import zipfile


HOST = "127.0.0.1"
PORT = 8772
BASE_API_URL = "https://api.crowdin.com/api/v2"
CROWDIN_PROJECT_SLUG = "operational-localization"
CROWDIN_TARGET_LANGUAGE = "en-US"
APITABLE_BASE_ORIGIN = "https://apitable.yottastudios.com"
APITABLE_DATASHEET_ID = "dst54Y1Wzwdm5sDeQ7"
APITABLE_READ_VIEW_ID = "viw8QB941rgSg"
APITABLE_WRITE_VIEW_ID = "viwJ88IVxdxo"
SCRIPT_DIR = Path(__file__).resolve().parent
APP_SUPPORT_DIR = Path.home() / "Library" / "Application Support" / "CrowdinAPITableCommentTool"
PACKAGED_STATE_FILE = APP_SUPPORT_DIR / "crowdin_apitable_comment_tool_state.json"
LEGACY_STATE_FILE = SCRIPT_DIR / "crowdin_apitable_comment_tool_state.json"
BATCH_FIELD_ASSIGN_LOCALIZER = "分配本地化（那三个人仅特殊情况才选择）"
BATCH_FIELD_REQUESTER = "分配需求者"
BATCH_FIELD_MESSAGE = "给本地化留言"
BATCH_FIELD_SEND_FLAG = "是否发送信息给本地化"
BATCH_FIELD_CHECK_LEAD = "检查提前量（min）"
SYNC_FIELD_SEARCH_KEY = "翻译需求（当日日期+翻译需求名字+需求人）"
READ_FIELD_REQUESTER = "需求人"
READ_FIELD_DEADLINE = "需求ddl（尽量提前1-3小时）"
READ_FIELD_PRIORITY = "优先级"
READ_FIELD_CROWDIN_FLAG = "是否用运营组专用crowdin"
READ_FIELD_TASK_LINK = "填写实际任务链接"
SYNC_FIELD_CHECKER = "检查者"
SYNC_FIELD_CHECKED = "是否检查完"
SYNC_FIELD_NOTE = "留言（会在检查完的消息里添加）"

try:
    import certifi
except Exception:
    certifi = None


DEFAULT_FORM = {
    "crowdin_token": "",
    "save_crowdin_token": True,
    "apitable_api_key": "",
    "save_apitable_api_key": True,
    "file_name": "",
    "crowdin_folder": "English Team",
    "extra_path_keyword": "",
    "extra_keyword": "",
    "apitable_search_column": "翻译需求（当日日期+翻译需求名字+需求人）",
    "apitable_comment_column": "完成版附件",
    "current_user": "仿生人",
    "sync_note": "",
}


def resolve_state_file() -> Path:
    if getattr(sys, "frozen", False):
        APP_SUPPORT_DIR.mkdir(parents=True, exist_ok=True)
        if not PACKAGED_STATE_FILE.exists() and LEGACY_STATE_FILE.exists():
            try:
                PACKAGED_STATE_FILE.write_text(LEGACY_STATE_FILE.read_text(encoding="utf-8"), encoding="utf-8")
            except Exception:
                pass
        return PACKAGED_STATE_FILE
    return LEGACY_STATE_FILE


STATE_FILE = resolve_state_file()

STATE: Dict[str, Any] = {
    "form": dict(DEFAULT_FORM),
    "jobs": {},
}
AUTO_EXIT_IDLE_SECONDS: Optional[float] = None
LAST_CLIENT_PING = 0.0
HAS_CLIENT_PING = False
SERVER_REF: Optional[ThreadingHTTPServer] = None
SHUTDOWN_STARTED = False


def mark_client_ping() -> None:
    global LAST_CLIENT_PING, HAS_CLIENT_PING
    LAST_CLIENT_PING = time.time()
    HAS_CLIENT_PING = True


def request_shutdown() -> None:
    global SHUTDOWN_STARTED
    if SHUTDOWN_STARTED:
        return
    SHUTDOWN_STARTED = True
    server = SERVER_REF
    if not server:
        return
    threading.Thread(target=server.shutdown, daemon=True).start()


def watchdog_auto_exit() -> None:
    if AUTO_EXIT_IDLE_SECONDS is None:
        return
    while True:
        time.sleep(1.0)
        if SHUTDOWN_STARTED:
            return
        if HAS_CLIENT_PING and time.time() - LAST_CLIENT_PING > AUTO_EXIT_IDLE_SECONDS:
            request_shutdown()
            return


def port_is_open(host: str, port: int) -> bool:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(0.5)
    try:
        return sock.connect_ex((host, port)) == 0
    finally:
        sock.close()


def ask_existing_instance_to_shutdown(host: str, port: int) -> None:
    try:
        req = urllib.request.Request(f"http://{host}:{port}/api/shutdown", method="POST")
        urllib.request.urlopen(req, timeout=0.8).read()
    except Exception:
        return
    for _ in range(20):
        if not port_is_open(host, port):
            return
        time.sleep(0.1)


def kill_stale_listener(port: int) -> None:
    try:
        result = subprocess.run(
            ["lsof", "-tiTCP:%d" % port, "-sTCP:LISTEN"],
            check=False,
            capture_output=True,
            text=True,
            timeout=1.0,
        )
    except Exception:
        return
    for raw_pid in result.stdout.splitlines():
        try:
            pid = int(raw_pid.strip())
        except ValueError:
            continue
        if pid == os.getpid():
            continue
        try:
            os.kill(pid, 15)
        except Exception:
            pass
    for _ in range(20):
        if not port_is_open(HOST, port):
            return
        time.sleep(0.1)


def create_server_with_fallback(host: str, preferred_port: int) -> Tuple[ThreadingHTTPServer, int]:
    ask_existing_instance_to_shutdown(host, preferred_port)
    if port_is_open(host, preferred_port):
        kill_stale_listener(preferred_port)
    last_error: Optional[Exception] = None
    for offset in range(20):
        port = preferred_port + offset
        try:
            return ThreadingHTTPServer((host, port), Handler), port
        except OSError as exc:
            last_error = exc
            if exc.errno in (48, 98):
                continue
            raise
    raise RuntimeError(f"无法启动本地服务，端口 {preferred_port}-{preferred_port + 19} 都不可用：{last_error}")

INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Crowdin → APITable 留言同步工具</title>
  <style>
    :root {
      --bg: #f4ede4;
      --bg-deep: #ece1d3;
      --panel: rgba(255, 251, 246, 0.86);
      --line: rgba(108, 90, 73, 0.16);
      --text: #1f1a17;
      --muted: #72675d;
      --primary: #145f59;
      --primary-strong: #0e4c48;
      --accent: #b86b35;
      --ok-soft: #e8f5ee;
      --danger-soft: #fbe9e6;
      --warning-soft: #fff4dc;
      --shadow-lg: 0 24px 60px rgba(67, 49, 29, 0.12);
      --shadow-md: 0 14px 32px rgba(67, 49, 29, 0.08);
      --shadow-sm: 0 8px 18px rgba(67, 49, 29, 0.05);
      --radius-xl: 28px;
      --radius-lg: 22px;
      --radius-md: 16px;
      --radius-sm: 12px;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Avenir Next", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
      color: var(--text);
      background:
        radial-gradient(920px 460px at 0% 0%, rgba(184, 107, 53, 0.18), transparent 56%),
        radial-gradient(780px 420px at 100% 0%, rgba(20, 95, 89, 0.16), transparent 52%),
        linear-gradient(180deg, #f7f1e9 0%, var(--bg) 58%, var(--bg-deep) 100%);
      padding: 30px 18px 44px;
    }
    .wrap {
      max-width: 1460px;
      margin: 0 auto;
    }
    .topbar {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(240px, 320px);
      gap: 18px;
      margin-bottom: 20px;
    }
    .hero-card,
    .meta-card,
    .surface,
    .workspace-shell,
    .feedback-shell {
      background: var(--panel);
      border: 1px solid rgba(255, 255, 255, 0.42);
      box-shadow: var(--shadow-lg);
      backdrop-filter: blur(14px);
    }
    .hero-card {
      border-radius: var(--radius-xl);
      padding: 28px 30px;
      position: relative;
      overflow: hidden;
    }
    .hero-card::after {
      content: "";
      position: absolute;
      width: 260px;
      height: 260px;
      right: -40px;
      bottom: -90px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(20, 95, 89, 0.14), transparent 70%);
    }
    .hero-heading {
      margin: 0;
      font-size: 37px;
      line-height: 1.08;
      letter-spacing: -0.02em;
      font-weight: 760;
      max-width: 720px;
    }
    .hero-sub {
      margin-top: 14px;
      max-width: 780px;
      color: var(--muted);
      font-size: 14px;
      line-height: 1.82;
    }
    .hero-notes {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-top: 18px;
    }
    .note-chip,
    .badge,
    .record-filter-tab {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 9px 13px;
      border: 1px solid var(--line);
      border-radius: 999px;
      background: rgba(255, 255, 255, 0.66);
      color: var(--muted);
      font-size: 12px;
      font-weight: 650;
      letter-spacing: 0.01em;
    }
    .meta-card {
      border-radius: var(--radius-xl);
      padding: 22px;
      display: flex;
      flex-direction: column;
      gap: 14px;
      align-items: flex-start;
      justify-content: space-between;
    }
    .meta-kicker {
      font-size: 12px;
      color: var(--accent);
      text-transform: uppercase;
      letter-spacing: 0.12em;
      font-weight: 700;
    }
    .meta-title {
      font-size: 14px;
      line-height: 1.75;
      color: var(--muted);
    }
    .meta-panel {
      border-radius: var(--radius-lg);
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.74);
      padding: 14px 16px;
    }
    .meta-panel .k,
    .subsurface-title,
    .card-label,
    .meta-box .k {
      font-size: 11px;
      color: var(--muted);
      text-transform: uppercase;
      letter-spacing: 0.1em;
      font-weight: 700;
      margin-bottom: 6px;
    }
    .meta-panel .v {
      font-size: 14px;
      line-height: 1.7;
      color: var(--text);
      word-break: break-word;
    }
    .dashboard {
      display: grid;
      grid-template-columns: minmax(320px, 390px) minmax(0, 1fr);
      grid-template-areas: "workspace context";
      gap: 20px;
      align-items: start;
    }
    main {
      display: grid;
      gap: 20px;
      align-self: start;
      grid-area: workspace;
      min-width: 0;
    }
    .surface,
    .workspace-shell,
    .feedback-shell {
      border-radius: 30px;
      padding: 22px;
    }
    .surface {
      min-height: calc(100vh - 240px);
      grid-area: context;
      min-width: 0;
    }
    .surface-title,
    .workspace-title,
    .feedback-title {
      margin: 0;
      font-size: 22px;
      font-weight: 750;
      letter-spacing: -0.02em;
    }
    .surface-copy,
    .workspace-copy,
    .feedback-copy {
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.76;
    }
    .global-stack,
    .context-stack,
    .workspace-stack {
      display: grid;
      gap: 16px;
    }
    .context-stack {
      grid-template-rows: auto minmax(0, 1fr);
    }
    .global-stack {
      grid-template-rows: auto minmax(520px, 1fr) auto auto;
    }
    .subsurface,
    .soft-panel,
    .console-card {
      background: linear-gradient(180deg, rgba(255,255,255,0.76), rgba(249,243,236,0.92));
      border: 1px solid rgba(108, 90, 73, 0.12);
      border-radius: 24px;
      padding: 16px;
      box-shadow: var(--shadow-sm);
      min-width: 0;
      overflow: hidden;
    }
    .subsurface.tight { padding: 14px; }
    .subsurface.compact {
      padding: 12px;
      border-radius: 18px;
    }
    .subsurface.task-panel {
      padding: 14px;
      min-height: 420px;
      display: grid;
      grid-template-rows: auto auto minmax(0, 1fr);
    }
    .global-controls,
    .mini-grid,
    .row,
    .inner-grid,
    .result-meta {
      display: grid;
      grid-template-columns: repeat(12, minmax(0, 1fr));
      gap: 12px;
    }
    .global-controls .block-5 { grid-column: span 5; }
    .global-controls .block-7 { grid-column: span 7; }
    .global-controls .block-12 { grid-column: span 12; }
    .global-controls.compact-config {
      grid-template-columns: 120px minmax(150px, 1fr) 176px;
      gap: 8px;
      align-items: end;
    }
    .global-controls.compact-config .block-5,
    .global-controls.compact-config .block-7,
    .global-controls.compact-config .block-12 {
      grid-column: auto;
    }
    .mini-grid .span-12 { grid-column: span 12; }
    .field {
      grid-column: span 6;
      display: flex;
      flex-direction: column;
      gap: 7px;
    }
    .field.full { grid-column: span 12; }
    .global-card,
    .meta-box {
      background: rgba(255, 255, 255, 0.74);
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 14px;
      box-shadow: var(--shadow-sm);
      min-width: 0;
      overflow: hidden;
    }
    .global-card.compact { padding: 15px 16px; }
    .global-card.slim {
      padding: 8px 10px;
      border-radius: 14px;
      box-shadow: none;
    }
    .global-card.slim input,
    .global-card.slim select {
      min-height: 36px;
      padding-top: 7px;
      padding-bottom: 7px;
    }
    .compact-config .global-actions {
      flex-wrap: nowrap;
      min-height: 55px;
      justify-content: center;
      gap: 8px;
    }
    .compact-config .global-actions button {
      min-width: 94px;
      padding: 9px 10px;
      white-space: nowrap;
    }
    .compact-config .inline-check {
      white-space: nowrap;
      font-size: 11px;
    }
    .focus-grid {
      display: grid;
      grid-template-columns: minmax(0, 0.8fr) minmax(0, 1.2fr);
      gap: 8px;
    }
    .global-value,
    .meta-box .v {
      font-size: 14px;
      line-height: 1.7;
      color: var(--text);
      word-break: break-word;
    }
    #translationPreview {
      max-height: 320px;
      overflow-y: auto;
      white-space: pre-wrap;
      font-family: ui-monospace, "SF Mono", Menlo, monospace;
      font-size: 12px;
      line-height: 1.65;
      padding: 8px 0;
    }
    .global-value.compact-value {
      font-size: 12px;
      line-height: 1.45;
      max-height: 36px;
      overflow: hidden;
    }
    label {
      font-size: 13px;
      color: var(--text);
      font-weight: 650;
      letter-spacing: 0.01em;
    }
    input, textarea, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: var(--radius-sm);
      padding: 12px 13px;
      font: inherit;
      color: var(--text);
      background: rgba(255, 255, 255, 0.92);
      outline: none;
      transition: border-color 0.16s ease, box-shadow 0.16s ease, background 0.16s ease;
    }
    select {
      appearance: none;
      background-image:
        linear-gradient(45deg, transparent 50%, var(--muted) 50%),
        linear-gradient(135deg, var(--muted) 50%, transparent 50%);
      background-position:
        calc(100% - 18px) calc(50% - 3px),
        calc(100% - 12px) calc(50% - 3px);
      background-size: 6px 6px, 6px 6px;
      background-repeat: no-repeat;
      padding-right: 34px;
    }
    input:focus, textarea:focus, select:focus {
      border-color: rgba(20, 95, 89, 0.48);
      box-shadow: 0 0 0 4px rgba(20, 95, 89, 0.08);
      background: #fff;
    }
    textarea {
      min-height: 150px;
      resize: vertical;
      line-height: 1.6;
      font-family: ui-monospace, "SF Mono", Menlo, monospace;
      background: rgba(252, 250, 247, 0.98);
    }
    .compact-textarea {
      min-height: 44px;
      height: 44px;
      max-height: 76px;
      line-height: 1.5;
      overflow-y: auto;
      resize: none;
    }
    .hint {
      color: var(--muted);
      font-size: 12px;
      line-height: 1.7;
    }
    .status.compact-status {
      padding: 12px 14px;
      border-radius: 16px;
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
    }
    .status.compact-status .status-title {
      margin: 0;
      white-space: nowrap;
      flex: 0 0 auto;
    }
    .status.compact-status .hint {
      line-height: 1.45;
      text-align: left;
      flex: 1 1 260px;
      min-width: 0;
      word-break: break-word;
    }
    .actions,
    .global-actions,
    .record-filter-bar,
    .feedback-head,
    .workspace-header {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }
    .feedback-head,
    .workspace-header {
      justify-content: space-between;
    }
    .workspace-config-trigger {
      width: 100%;
      justify-content: center;
    }
    button {
      border: 1px solid transparent;
      border-radius: 14px;
      padding: 11px 16px;
      font: inherit;
      font-weight: 700;
      letter-spacing: 0.01em;
      cursor: pointer;
      background: linear-gradient(135deg, var(--primary), var(--primary-strong));
      color: #fff;
      box-shadow: 0 12px 22px rgba(20, 95, 89, 0.18);
      transition: transform 0.16s ease, box-shadow 0.16s ease, background 0.16s ease;
    }
    button:hover {
      transform: translateY(-1px);
      box-shadow: 0 16px 28px rgba(20, 95, 89, 0.22);
    }
    button.secondary,
    .record-filter-tab {
      background: rgba(255, 255, 255, 0.76);
      color: var(--text);
      border-color: var(--line);
      box-shadow: none;
    }
    button.secondary:hover {
      background: rgba(255, 255, 255, 0.96);
      box-shadow: var(--shadow-sm);
    }
    button[disabled] { opacity: 0.6; cursor: wait; }
    .record-filter-tab.active,
    .workspace-tab.active {
      background: linear-gradient(135deg, rgba(20,95,89,0.98), rgba(13,79,74,0.98));
      color: #f7fffd;
      border-color: transparent;
      box-shadow: 0 12px 26px rgba(20, 95, 89, 0.16);
    }
    .inline-check {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      color: var(--muted);
      font-size: 12px;
      font-weight: 600;
    }
    .inline-check input {
      width: auto;
      accent-color: var(--primary);
    }
    .status {
      border-radius: 18px;
      padding: 16px;
      border: 1px solid var(--line);
      background: rgba(255, 255, 255, 0.74);
      box-shadow: var(--shadow-sm);
    }
    .status.running {
      background: var(--warning-soft);
      border-color: rgba(184, 107, 53, 0.24);
    }
    .status.success {
      background: var(--ok-soft);
      border-color: rgba(22, 106, 69, 0.22);
    }
    .status.error {
      background: var(--danger-soft);
      border-color: rgba(169, 54, 40, 0.18);
    }
    .status-title {
      font-weight: 760;
      margin-bottom: 10px;
    }
    .record-list {
      display: grid;
      gap: 12px;
      max-height: none;
      overflow: auto;
      padding-right: 4px;
      align-content: start;
    }
    .record-card-item {
      border: 1px solid rgba(108, 90, 73, 0.14);
      border-radius: 22px;
      background: linear-gradient(180deg, rgba(255,255,255,0.96), rgba(249,243,236,0.96));
      padding: 16px;
      cursor: pointer;
      box-shadow: var(--shadow-sm);
      transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease;
    }
    .record-card-item:hover {
      transform: translateY(-2px);
      box-shadow: 0 18px 36px rgba(67, 49, 29, 0.10);
    }
    .record-card-item.selected {
      border-color: rgba(20, 95, 89, 0.34);
      background: linear-gradient(180deg, #eef8f6, #f9fffd);
      box-shadow: 0 20px 42px rgba(20, 95, 89, 0.14);
    }
    .record-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: start;
      margin-bottom: 12px;
    }
    .record-id {
      font-size: 12px;
      color: var(--muted);
      margin-bottom: 6px;
    }
    .record-title {
      font-size: 17px;
      font-weight: 750;
      line-height: 1.45;
      color: var(--text);
      word-break: break-word;
    }
    .record-badges {
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .record-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 10px;
    }
    .record-meta {
      background: rgba(255,255,255,0.70);
      border: 1px solid rgba(108, 90, 73, 0.10);
      border-radius: 16px;
      padding: 10px 12px;
      min-height: 58px;
      min-width: 0;
      overflow: hidden;
    }
    .record-meta .meta-label {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.1em;
      color: var(--accent);
      margin-bottom: 6px;
      font-weight: 700;
    }
    .record-meta .meta-value {
      font-size: 13px;
      line-height: 1.6;
      color: var(--text);
      word-break: break-word;
    }
    .option-note {
      display: none;
    }
    .workspace-nav {
      display: grid;
      grid-template-columns: 1fr;
      gap: 10px;
      width: 100%;
    }
    .workspace-tab {
      width: 100%;
      text-align: left;
      display: grid;
      gap: 4px;
      border-radius: 18px;
      padding: 14px 15px;
    }
    .workspace-tab.secondary {
      color: var(--muted);
      text-align: left;
      min-height: 0;
      box-shadow: none;
    }
    .workspace-tab .tab-label {
      display: block;
      font-size: 15px;
      font-weight: 740;
      margin-bottom: 4px;
      color: inherit;
      white-space: normal;
    }
    .workspace-tab .tab-copy {
      display: block;
      font-size: 12px;
      line-height: 1.6;
      color: inherit;
      opacity: 0.92;
      white-space: normal;
    }
    .workspace-shell .row,
    .workspace-shell .inner-grid,
    .workspace-shell .result-meta {
      grid-template-columns: 1fr;
    }
    .workspace-shell .field,
    .workspace-shell .field.full,
    .workspace-shell .inner-grid > section,
    .workspace-shell .result-meta > .meta-box {
      grid-column: 1 / -1 !important;
    }
    .workspace-panel { display: none; }
    .workspace-panel.active { display: block; }
    .console-grid {
      display: grid;
      grid-template-columns: minmax(0, 0.38fr) minmax(0, 0.62fr);
      gap: 16px;
    }
    .console-card.dark {
      background: transparent;
      border: 0;
      padding: 0;
      box-shadow: none;
      overflow: visible;
    }
    .log {
      background: linear-gradient(180deg, rgba(21, 31, 41, 0.98), rgba(15, 23, 32, 0.98));
      color: #d9e4ef;
      border-radius: 20px;
      padding: 16px;
      min-height: 400px;
      max-height: 600px;
      white-space: pre-wrap;
      line-height: 1.65;
      font-size: 12px;
      font-family: ui-monospace, "SF Mono", Menlo, monospace;
      overflow: auto;
      border: 1px solid rgba(255, 255, 255, 0.06);
    }
    .meta-box { grid-column: span 4; }
    .muted-block {
      padding: 14px 15px;
      border-radius: 18px;
      background: rgba(249, 241, 232, 0.84);
      border: 1px solid rgba(108, 90, 73, 0.10);
      margin-top: 2px;
    }
    .record-card-item,
    .status,
    .workspace-shell,
    .surface {
      min-width: 0;
      overflow: hidden;
    }
    .feedback-shell {
      min-width: 0;
      overflow: visible;
    }
    .surface-toolbar {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      justify-content: flex-end;
    }
    .modal {
      position: fixed;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      padding: 24px;
      background: rgba(31, 26, 23, 0.38);
      z-index: 20;
    }
    .modal.open { display: flex; }
    .modal-panel {
      width: min(780px, 100%);
      max-height: min(86vh, 960px);
      overflow: auto;
      border-radius: 30px;
      padding: 22px;
      background: linear-gradient(180deg, rgba(255,255,255,0.97), rgba(249,243,236,0.96));
      border: 1px solid rgba(255, 255, 255, 0.58);
      box-shadow: var(--shadow-lg);
    }
    .modal-head {
      display: flex;
      align-items: flex-start;
      justify-content: space-between;
      gap: 14px;
      margin-bottom: 16px;
    }
    .modal-copy {
      margin-top: 8px;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.75;
    }
    .modal-close {
      min-width: 84px;
      box-shadow: none;
    }
    .modal-grid {
      display: grid;
      gap: 14px;
    }
    .modal-grid.two {
      grid-template-columns: repeat(2, minmax(0, 1fr));
    }
    @media (max-width: 920px) {
      .topbar,
      .dashboard,
      .console-grid {
        grid-template-columns: 1fr;
        grid-template-areas:
          "workspace"
          "context";
      }
      main,
      .surface {
        grid-area: auto;
      }
      .field,
      .meta-box,
      .global-controls .block-5,
      .global-controls .block-7,
      .global-controls .block-12 {
        grid-column: span 12;
      }
      .global-controls.compact-config,
      .focus-grid {
        grid-template-columns: 1fr;
      }
      .global-controls.compact-config .block-5,
      .global-controls.compact-config .block-7,
      .global-controls.compact-config .block-12 {
        grid-column: span 12;
      }
      .record-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .hero-heading { font-size: 29px; }
      .modal-grid.two { grid-template-columns: 1fr; }
    }
    @media (max-width: 720px) {
      body { padding: 18px 12px 28px; }
      .record-grid,
      .result-meta {
        grid-template-columns: 1fr;
      }
      .meta-box { grid-column: span 12; }
      .modal {
        padding: 14px;
      }
      .modal-head {
        flex-direction: column;
      }
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="topbar">
      <section class="hero-card">
        <h1 class="hero-heading">Crowdin → APITable 留言同步工具</h1>
        <div class="hero-sub">
          把“选任务、分配本地化、同步完成版附件、回写检查信息”整理到同一个工作台里。右侧放大保留全局上下文和整表卡片，左侧收拢执行参数，减少来回切换造成的中断。
        </div>
        <div class="hero-notes">
          <div class="note-chip">先读表并选中目标记录</div>
          <div class="note-chip">两个工作区共享同一份上下文</div>
          <div class="note-chip">同步后自动补检查与留言</div>
        </div>
      </section>

      <aside class="meta-card">
        <div>
          <div class="meta-kicker">Workspace Context</div>
          <div class="meta-title">固定视图、当前使用人和 APITable Key 都收进同一个入口里了，需要时再展开查看。</div>
        </div>
        <div class="workspace-nav">
          <button type="button" id="loadTableBtnTop" class="secondary workspace-config-trigger">读取整张表</button>
          <button type="button" id="openWorkspaceConfigBtn" class="secondary workspace-config-trigger">打开工作区设置</button>
        </div>
      </aside>
    </div>

    <div id="workspaceConfigModal" class="modal" aria-hidden="true">
      <div class="modal-panel">
        <div class="modal-head">
          <div>
            <h2 class="surface-title">工作区设置</h2>
            <div class="modal-copy">这里集中放固定读写视图和共享参数。平时默认收起，不占用主工作区。</div>
          </div>
          <button type="button" id="closeWorkspaceConfigBtn" class="secondary modal-close">关闭</button>
        </div>

        <div class="modal-grid two">
          <div class="meta-panel">
            <div class="k">读取视图</div>
            <div class="v">dst54Y1Wzwdm5sDeQ7 / viw8QB941rgSg</div>
          </div>
          <div class="meta-panel">
            <div class="k">写入视图</div>
            <div class="v">dst54Y1Wzwdm5sDeQ7 / viwJ88IVxdxo</div>
          </div>
        </div>

        <div class="subsurface compact" style="margin-top: 14px;">
          <div class="subsurface-title">共享条件</div>
          <div class="global-controls compact-config">
            <div class="block-5 global-card slim">
              <div class="card-label">当前使用人</div>
              <select id="current_user">
                <option value="仿生人">仿生人</option>
                <option value="石上">石上</option>
                <option value="德古拉">德古拉</option>
              </select>
            </div>
            <div class="block-7 global-card slim">
              <div class="card-label">APITable API Key</div>
              <input id="apitable_api_key" type="password" placeholder="输入 APITable API Key">
            </div>
            <div class="block-12 global-card slim">
              <div class="global-actions">
                <label class="inline-check"><input id="save_apitable_api_key" type="checkbox"> 保存 API Key</label>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div class="dashboard">
      <aside class="surface">
        <div class="context-stack">
          <div>
            <h2 class="surface-title">全局上下文</h2>
            <div class="surface-copy">这里保留整张表、当前选中行和搜索关键词，作为整个页面最主要的工作区域。左侧参数改动后，会直接作用到这里选中的记录。</div>
          </div>

          <div class="global-stack">
            <div class="surface-toolbar">
              <button type="button" id="loadTableBtnInline" class="secondary">读取整张表</button>
              <button type="button" id="openWorkspaceConfigInlineBtn" class="secondary">查看工作区设置</button>
            </div>

            <div class="subsurface tight task-panel">
              <div class="subsurface-title">任务卡片</div>
              <div class="record-filter-bar">
                <button type="button" id="recordFilterFalse" class="record-filter-tab active">未发送</button>
                <button type="button" id="recordFilterTrue" class="record-filter-tab">已发送</button>
              </div>
              <div id="batchTableBody" class="record-list">
                <div class="record-card-item">
                  <div class="record-id">尚未读取表格</div>
                  <div class="record-title">读取视图后，这里会按任务卡片展示数据。</div>
                </div>
              </div>
            </div>

            <div class="subsurface compact">
              <div class="subsurface-title">当前焦点</div>
              <div class="focus-grid">
                <div class="global-card slim">
                  <div class="card-label">当前选中行</div>
                  <div id="selectedBatchRow" class="global-value compact-value">尚未选择。</div>
                </div>
                <div class="global-card slim">
                  <div class="card-label">Crowdin 搜索关键词</div>
                  <div id="selectedSearchKeyword" class="global-value compact-value">尚未选择。</div>
                </div>
              </div>
            </div>

            <div id="batchStatus" class="status compact-status">
              <div class="status-title">等待读取整表</div>
              <div class="hint">点“读取整张表”后，这里会显示字段选项和应用结果。</div>
            </div>
          </div>
        </div>
      </aside>

      <main>
        <section class="workspace-shell">
          <div class="workspace-header">
            <div>
              <h2 class="workspace-title">执行工作区</h2>
              <div class="workspace-copy">这里专门放执行参数。先在右侧选中目标任务，再回来切换这两个动作面板完成写入和同步。</div>
            </div>
            <div class="workspace-nav">
              <button type="button" class="workspace-tab secondary active" data-panel="batchPanel">
                <span class="tab-label">表格批量操作</span>
                <span class="tab-copy">写入分配、留言、检查提前量与发送状态。</span>
              </button>
              <button type="button" class="workspace-tab secondary" data-panel="syncPanel">
                <span class="tab-label">Crowdin 附件同步</span>
                <span class="tab-copy">按翻译需求搜完成版并回填附件与检查信息。</span>
              </button>
            </div>
          </div>

          <section class="workspace-panel active" id="batchPanel">
            <div class="soft-panel">
              <div class="hint" style="margin-bottom: 14px;">
                这里负责把当前填写的模板直接应用到左侧选中的那一行。前四项立即更新，最后的“是否发送信息给本地化”会延迟 3 秒再勾选。
              </div>
              <div class="actions" style="margin-bottom: 14px;">
                <button id="applyBatchTemplateBtn">应用到选中行</button>
              </div>
              <div class="row">
                <div class="field">
                  <label>分配本地化</label>
                  <select id="batch_assign_localizer"></select>
                  <div id="batch_assign_localizer_note" class="option-note">读取整张表后显示选项。</div>
                </div>
                <div class="field">
                  <label>分配需求者</label>
                  <select id="batch_requester"></select>
                  <div id="batch_requester_note" class="option-note">读取整张表后显示选项。</div>
                </div>
                <div class="field full">
                  <label>给本地化留言</label>
                  <textarea id="batch_message" class="compact-textarea" placeholder="填写要写入“给本地化留言”的内容"></textarea>
                  <div id="batch_message_note" class="option-note">如果这一列已有常见内容，也会在这里提示。</div>
                </div>
                <div class="field">
                  <label>检查提前量（min）</label>
                  <select id="batch_check_lead"></select>
                  <div id="batch_check_lead_note" class="option-note">读取整张表后显示选项，默认优先选择 30。</div>
                </div>
                <div class="field">
                  <label class="inline-check"><input id="batch_send_flag" type="checkbox" checked> 是否发送信息给本地化</label>
                  <div id="batch_send_flag_note" class="option-note">勾选动作会比前三列晚 3 秒执行。</div>
                </div>
              </div>
            </div>
          </section>

          <section class="workspace-panel" id="syncPanel">
            <div class="workspace-stack">
              <div class="soft-panel">
                <div class="hint" style="margin-bottom: 14px;">
                  默认会直接使用左侧当前选中行里的翻译需求作为搜索关键词，你也可以在这里微调路径关键词和附加关键词。
                </div>
                <div class="inner-grid">
                  <section class="subsurface tight" style="grid-column: span 6;">
                    <div class="subsurface-title">Crowdin 凭证</div>
                    <div class="row">
                      <div class="field full">
                        <label>Crowdin API Token</label>
                        <input id="crowdin_token" type="password" placeholder="输入 Crowdin Personal Access Token">
                      </div>
                      <div class="field full">
                        <label class="inline-check"><input id="save_crowdin_token" type="checkbox"> 保存 Crowdin Token 到本地</label>
                      </div>
                    </div>
                  </section>

                  <section class="subsurface tight" style="grid-column: span 6;">
                    <div class="subsurface-title">匹配条件</div>
                    <div class="row">
                      <div class="field full">
                        <label>文件名或主关键词</label>
                        <input id="file_name" type="text" placeholder="例如 weapons.json 或 weapons">
                      </div>
                      <div class="field">
                        <label>Crowdin 默认路径</label>
                        <input id="crowdin_folder" type="text" placeholder="English Team">
                      </div>
                      <div class="field">
                        <label>附加路径关键词</label>
                        <input id="extra_path_keyword" type="text" placeholder="可选，例如 activity/mail">
                      </div>
                      <div class="field full">
                        <label>附加文件关键词</label>
                        <input id="extra_keyword" type="text" placeholder="可选，进一步缩小匹配范围">
                      </div>
                    </div>
                  </section>

                  <section class="subsurface tight" style="grid-column: 1 / -1;">
                    <div class="subsurface-title">APITable 回写规则</div>
                    <div class="row">
                      <div class="field">
                        <label>搜索列名</label>
                        <input id="apitable_search_column" type="text" placeholder="默认：翻译需求（当日日期+翻译需求名字+需求人）">
                      </div>
                      <div class="field">
                        <label>写入列名</label>
                        <input id="apitable_comment_column" type="text" placeholder="默认：完成版附件">
                      </div>
                      <div class="field full">
                        <label>检查完成留言</label>
                        <textarea id="sync_note" placeholder="同步完成后，自动写入“留言（会在检查完的消息里添加）”"></textarea>
                      </div>
                    </div>
                    <div class="muted-block">
                      <div class="hint">
                        命中目标行后，会自动上传完成版附件，并把当前使用人写入“检查者”、勾选“是否检查完”，再写入上面的检查完成留言。
                      </div>
                    </div>
                    <div class="actions" style="margin-top: 14px;">
                      <button id="runBtn">开始同步</button>
                      <button id="saveBtn" class="secondary">仅保存本地配置</button>
                    </div>
                  </section>
                </div>
              </div>
            </div>
          </section>
        </section>
      </main>
    </div>

    <section class="feedback-shell">
      <div class="feedback-head">
        <div>
          <h2 class="feedback-title">实时反馈</h2>
          <div class="feedback-copy">同步任务启动后，这里会持续显示状态、日志和结果摘要，方便你边看过程边复制翻译内容。</div>
        </div>
        <div class="actions" style="margin: 0;">
          <button id="copyBtn" class="secondary">复制翻译结果</button>
        </div>
      </div>
      <div class="console-grid">
        <div class="console-card">
          <div id="statusBox" class="status">
            <div class="status-title">等待执行</div>
            <div class="hint">点击“开始同步”后，这里会显示执行状态。</div>
          </div>
          <div class="result-meta">
            <div class="meta-box">
              <div class="k">匹配文件</div>
              <div id="matchedFile" class="v">-</div>
            </div>
            <div class="meta-box">
              <div class="k">匹配记录</div>
              <div id="matchedRecord" class="v">-</div>
            </div>
            <div class="meta-box">
              <div class="k">翻译预览</div>
              <div id="translationPreview" class="v">-</div>
            </div>
          </div>
        </div>
        <div class="console-card dark">
          <div id="logBox" class="log">尚未开始。</div>
        </div>
      </div>
    </section>

  </div>

  <script>
    let currentJobId = "";
    let pollTimer = null;
    let lastTranslation = "";
    let batchTablePayload = null;
    let selectedBatchRecordId = "";
    let batchRecordFilter = "false";

    function byId(id) { return document.getElementById(id); }

    function setWorkspaceConfigModal(open) {
      const modal = byId("workspaceConfigModal");
      if (!modal) return;
      modal.classList.toggle("open", !!open);
      modal.setAttribute("aria-hidden", open ? "false" : "true");
      document.body.style.overflow = open ? "hidden" : "";
    }

    function getFormPayload() {
      return {
        crowdin_token: byId("crowdin_token").value.trim(),
        save_crowdin_token: byId("save_crowdin_token").checked,
        apitable_api_key: byId("apitable_api_key").value.trim(),
        save_apitable_api_key: byId("save_apitable_api_key").checked,
        current_user: byId("current_user").value.trim(),
        file_name: byId("file_name").value.trim(),
        crowdin_folder: byId("crowdin_folder").value.trim(),
        extra_path_keyword: byId("extra_path_keyword").value.trim(),
        extra_keyword: byId("extra_keyword").value.trim(),
        apitable_search_column: byId("apitable_search_column").value.trim(),
        apitable_comment_column: byId("apitable_comment_column").value.trim(),
        sync_note: byId("sync_note").value
      };
    }

    function getBatchApplyPayload() {
      return {
        assign_localizer: readOptionLabel("batch_assign_localizer"),
        requester: readOptionLabel("batch_requester"),
        message: byId("batch_message").value,
        check_lead: readOptionLabel("batch_check_lead"),
        send_flag: byId("batch_send_flag").checked
      };
    }

    function readOptionValue(id) {
      const el = byId(id);
      if (!el || !el.value) return "";
      try {
        return JSON.parse(el.value);
      } catch (_) {
        return el.value;
      }
    }

    function readOptionLabel(id) {
      const el = byId(id);
      if (!el) return "";
      const opt = el.options[el.selectedIndex];
      if (!opt || !opt.value) return "";
      return opt.dataset.label || opt.textContent || "";
    }

    function setFormValues(data) {
      const form = Object.assign({}, data || {});
      Object.keys(form).forEach((key) => {
        const el = byId(key);
        if (!el) return;
        if (el.type === "checkbox") {
          el.checked = !!form[key];
        } else {
          el.value = form[key] == null ? "" : String(form[key]);
        }
      });
    }

    function setRunning(running) {
      byId("runBtn").disabled = running;
      byId("saveBtn").disabled = running;
    }

    function updateStatus(job) {
      const box = byId("statusBox");
      const title = job.status === "running"
        ? "正在执行"
        : job.status === "success"
          ? "执行完成"
          : job.status === "error"
            ? "执行失败"
            : "等待执行";
      box.className = "status " + (job.status || "");
      box.innerHTML = "<div class='status-title'>" + title + "</div>"
        + "<div class='hint'>" + escapeHtml(job.message || "暂无说明") + "</div>";

      byId("logBox").textContent = (job.logs || []).join("\n") || "暂无日志。";
      byId("matchedFile").textContent = job.result && job.result.matched_file_path ? job.result.matched_file_path : "-";
      byId("matchedRecord").textContent = job.result && job.result.record_id ? (job.result.record_id + " / " + (job.result.search_column || "")) : "-";
      lastTranslation = job.result && job.result.translation_text ? job.result.translation_text : "";
      byId("translationPreview").textContent = lastTranslation || "-";
    }

    function escapeHtml(text) {
      return String(text || "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;");
    }

    function isTruthySendFlag(value) {
      const text = String(value == null ? "" : value).trim().toLowerCase();
      return text === "true" || text === "yes" || text === "1" || text === "是";
    }

    function renderRecordFilterTabs(rows) {
      const falseCount = (rows || []).filter((row) => !isTruthySendFlag((row.summary || {}).send_flag)).length;
      const trueCount = (rows || []).filter((row) => isTruthySendFlag((row.summary || {}).send_flag)).length;
      byId("recordFilterFalse").textContent = "未发送 " + falseCount;
      byId("recordFilterTrue").textContent = "已发送 " + trueCount;
      byId("recordFilterFalse").classList.toggle("active", batchRecordFilter === "false");
      byId("recordFilterTrue").classList.toggle("active", batchRecordFilter === "true");
    }

    async function loadState() {
      const resp = await fetch("/api/state");
      const data = await resp.json();
      setFormValues(data.form || {});
      syncRequesterWithCurrentUser(false);
      resizeBatchMessageArea();
    }

    function setBatchStatus(kind, title, message) {
      const box = byId("batchStatus");
      box.className = "status compact-status " + (kind || "");
      box.innerHTML = "<div class='status-title'>" + escapeHtml(title) + "</div><div class='hint'>" + escapeHtml(message || "") + "</div>";
    }

    function resizeBatchMessageArea() {
      const el = byId("batch_message");
      if (!el) return;
      el.style.height = "44px";
      const nextHeight = Math.min(Math.max(el.scrollHeight, 44), 76);
      el.style.height = nextHeight + "px";
    }

    function renderOptionSelect(selectId, noteId, options, placeholder) {
      const select = byId(selectId);
      const note = byId(noteId);
      select.innerHTML = "";
      const first = document.createElement("option");
      first.value = "";
      first.textContent = placeholder;
      select.appendChild(first);
      const seenLabels = new Set();
      const filtered = (options || []).filter((item) => {
        const label = String(item.label || "");
        if (!label) return false;
        if (/^opt[a-z0-9]+$/i.test(label)) return false;
        if (seenLabels.has(label)) return false;
        seenLabels.add(label);
        return true;
      });
      filtered.forEach((item) => {
        const opt = document.createElement("option");
        opt.value = JSON.stringify(item.value);
        opt.textContent = item.label || String(item.value || "");
        opt.dataset.label = item.label || String(item.value || "");
        select.appendChild(opt);
      });
      note.textContent = filtered.length
        ? "候选项：" + filtered.slice(0, 8).map((item) => item.label).join(" / ")
        : "未读取到明确选项，可先读取整表后再继续填写。";
    }

    function setSelectByRawValue(id, rawValue) {
      const select = byId(id);
      const text = rawValue == null ? "" : String(rawValue);
      let target = Array.from(select.options).find((opt) => (opt.dataset.label || opt.textContent || "") === text);
      if (!target && text) {
        const opt = document.createElement("option");
        opt.value = JSON.stringify(text);
        opt.textContent = text;
        opt.dataset.label = text;
        select.appendChild(opt);
        target = opt;
      }
      select.value = target ? target.value : "";
    }

    function collectEditorState() {
      return {
        batchValues: getBatchApplyPayload(),
        selectedRecordId: selectedBatchRecordId
      };
    }

    function restoreSelectedBatchRow() {
      const body = byId("batchTableBody");
      Array.from(body.querySelectorAll(".record-card-item")).forEach((item) => {
        item.classList.toggle("selected", item.getAttribute("data-record-id") === selectedBatchRecordId);
      });
      byId("selectedBatchRow").textContent = selectedBatchRecordId || "尚未选择。";
      if (!selectedBatchRecordId) {
        byId("selectedSearchKeyword").textContent = "尚未选择。";
      }
    }

    function syncRequesterWithCurrentUser(forceApply = true) {
      const currentUser = byId("current_user").value.trim();
      if (!currentUser) return;
      if (!forceApply && byId("batch_requester").value) return;
      setSelectByRawValue("batch_requester", currentUser);
    }

    function renderBatchTable(payload) {
      batchTablePayload = payload;
      const body = byId("batchTableBody");
      const allRows = payload.rows || [];
      renderRecordFilterTabs(allRows);
      const rows = allRows.filter((row) => {
        const isTrue = isTruthySendFlag((row.summary || {}).send_flag);
        return batchRecordFilter === "true" ? isTrue : !isTrue;
      });
      if (!rows.length) {
        body.innerHTML = "<div class='record-card-item'><div class='record-id'>当前分组没有记录</div><div class='record-title'>可以切换到另一个发送分组继续查看。</div></div>";
        return;
      }
      body.innerHTML = rows.map((row) => {
        const summary = row.summary || {};
        return "<article class='record-card-item' data-record-id='" + escapeHtml(row.recordId) + "'>"
          + "<div class='record-head'>"
          +   "<div>"
          +     "<div class='record-id'>" + escapeHtml(row.recordId) + "</div>"
          +     "<div class='record-title'>" + escapeHtml(summary.title || "未命名翻译需求") + "</div>"
          +   "</div>"
          +   "<div class='record-badges'>"
          +     (summary.priority ? "<span class='badge'>优先级 " + escapeHtml(summary.priority) + "</span>" : "")
          +     (summary.crowdin_flag ? "<span class='badge'>Crowdin " + escapeHtml(summary.crowdin_flag) + "</span>" : "")
          +     (summary.send_flag ? "<span class='badge'>发送 " + escapeHtml(summary.send_flag) + "</span>" : "")
          +   "</div>"
          + "</div>"
          + "<div class='record-grid'>"
          +   "<div class='record-meta'><div class='meta-label'>需求人</div><div class='meta-value'>" + escapeHtml(summary.requester || "-") + "</div></div>"
          +   "<div class='record-meta'><div class='meta-label'>需求 DDL</div><div class='meta-value'>" + escapeHtml(summary.deadline || "-") + "</div></div>"
          +   "<div class='record-meta'><div class='meta-label'>分配本地化</div><div class='meta-value'>" + escapeHtml(summary.localizer || "-") + "</div></div>"
          +   "<div class='record-meta'><div class='meta-label'>分配需求者</div><div class='meta-value'>" + escapeHtml(summary.owner || "-") + "</div></div>"
          +   "<div class='record-meta'><div class='meta-label'>检查提前量</div><div class='meta-value'>" + escapeHtml(summary.check_lead || "-") + "</div></div>"
          +   "<div class='record-meta'><div class='meta-label'>任务链接</div><div class='meta-value'>" + escapeHtml(summary.task_link || "-") + "</div></div>"
          + "</div>"
          + "</article>";
      }).join("");
      Array.from(body.querySelectorAll(".record-card-item")).forEach((card) => {
        card.addEventListener("click", () => {
          Array.from(body.querySelectorAll(".record-card-item")).forEach((item) => item.classList.remove("selected"));
          card.classList.add("selected");
          selectedBatchRecordId = card.getAttribute("data-record-id") || "";
          byId("selectedBatchRow").textContent = selectedBatchRecordId;
          const row = allRows.find((item) => item.recordId === selectedBatchRecordId);
          if (row && row.cells && row.cells["翻译需求（当日日期+翻译需求名字+需求人）"]) {
            const keyword = row.cells["翻译需求（当日日期+翻译需求名字+需求人）"];
            byId("file_name").value = keyword;
            byId("selectedSearchKeyword").textContent = keyword;
          }
        });
      });
    }

    async function loadBatchTable(preserveEditor = true) {
      const snapshot = preserveEditor ? collectEditorState() : null;
      setBatchStatus("running", "正在读取整张表", "请稍候…");
      try {
        const resp = await fetch("/api/batch-table/load", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ api_key: byId("apitable_api_key").value.trim() })
        });
        const data = await resp.json();
        if (!data.ok) throw new Error(data.error || "读取失败");
        const localizerMeta = data.table.meta["分配本地化（那三个人仅特殊情况才选择）"] || {};
        const requesterMeta = data.table.meta["分配需求者"] || {};
        const messageMeta = data.table.meta["给本地化留言"] || {};
        const checkLeadMeta = data.table.meta["检查提前量（min）"] || {};
        renderOptionSelect("batch_assign_localizer", "batch_assign_localizer_note", localizerMeta.options || [], "请选择分配本地化");
        renderOptionSelect("batch_requester", "batch_requester_note", requesterMeta.options || [], "请选择分配需求者");
        renderOptionSelect("batch_check_lead", "batch_check_lead_note", checkLeadMeta.options || [], "请选择检查提前量");
        byId("batch_message_note").textContent = (messageMeta.options || []).slice(0, 6).map((item) => item.label).join(" / ") || "这一列暂无常见候选。";
        renderBatchTable(data.table);
        if (snapshot) {
          const batchValues = snapshot.batchValues || {};
          byId("batch_message").value = batchValues.message || "";
          resizeBatchMessageArea();
          byId("batch_send_flag").checked = batchValues.send_flag !== false;
          setSelectByRawValue("batch_assign_localizer", batchValues.assign_localizer || "");
          setSelectByRawValue("batch_requester", batchValues.requester || "");
          setSelectByRawValue("batch_check_lead", batchValues.check_lead || "30");
          selectedBatchRecordId = snapshot.selectedRecordId || "";
          restoreSelectedBatchRow();
        }
        if (!byId("batch_check_lead").value) {
          setSelectByRawValue("batch_check_lead", "30");
        }
        if (!byId("batch_send_flag").checked) {
          byId("batch_send_flag").checked = true;
        }
        syncRequesterWithCurrentUser(false);
        setBatchStatus("success", "整表读取完成", "现在可以选择行并直接应用到目标记录。");
      } catch (err) {
        setBatchStatus("error", "读取失败", err.message);
        alert("读取整张表失败：" + err.message);
      }
    }

    async function applyBatchTemplate() {
      if (!selectedBatchRecordId) {
        alert("请先从表格里选中一行。");
        return;
      }
      setBatchStatus("running", "正在写入 APITable", "前三列会立即写入，最后一个勾选会延迟 3 秒。");
      try {
        const resp = await fetch("/api/batch-apply", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            api_key: byId("apitable_api_key").value.trim(),
            record_id: selectedBatchRecordId,
            template: getBatchApplyPayload()
          })
        });
        const data = await resp.json();
        if (!data.ok) throw new Error(data.error || "写入失败");
        let detail = data.message || "内容已写入选中行。";
        if (data.debug) {
          detail += "\n\n调试信息：\n" + JSON.stringify(data.debug, null, 2);
        }
        setBatchStatus("success", "写入完成", detail);
        await loadBatchTable();
      } catch (err) {
        setBatchStatus("error", "写入失败", err.message);
        alert("写入失败：" + err.message);
      }
    }

    async function saveConfig() {
      const resp = await fetch("/api/save-config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(getFormPayload())
      });
      return resp.json();
    }

    async function startSync() {
      setRunning(true);
      byId("statusBox").className = "status running";
      byId("statusBox").innerHTML = "<div class='status-title'>正在提交任务</div><div class='hint'>请稍候…</div>";
      try {
        const resp = await fetch("/api/start-sync", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(getFormPayload())
        });
        const data = await resp.json();
        if (!data.ok) {
          throw new Error(data.error || "启动失败");
        }
        currentJobId = data.job_id;
        pollJob();
      } catch (err) {
        setRunning(false);
        alert("启动失败：" + err.message);
      }
    }

    async function pollJob() {
      if (!currentJobId) return;
      try {
        const resp = await fetch("/api/job/" + currentJobId);
        const job = await resp.json();
        updateStatus(job);
        if (job.status === "running" || job.status === "pending") {
          pollTimer = setTimeout(pollJob, 900);
          return;
        }
        setRunning(false);
        if (job.status === "error") {
          if ((job.error_code || "") === "crowdin_not_found") {
            alert("未搜到对应文件，请检查文件名、路径关键词或附加关键词。");
          } else {
            alert(job.message || "执行失败");
          }
        }
      } catch (err) {
        setRunning(false);
        alert("读取进度失败：" + err.message);
      }
    }

    async function copyTranslation() {
      if (!lastTranslation) {
        alert("当前没有可复制的翻译结果。");
        return;
      }
      await navigator.clipboard.writeText(lastTranslation);
      alert("翻译结果已复制。");
    }

    function sendHeartbeat() {
      fetch("/api/ping", { method: "POST", keepalive: true }).catch(() => {});
    }

    byId("runBtn").addEventListener("click", startSync);
    byId("saveBtn").addEventListener("click", async () => {
      try {
        const data = await saveConfig();
        if (!data.ok) throw new Error(data.error || "保存失败");
        alert("本地配置已保存。");
      } catch (err) {
        alert("保存失败：" + err.message);
      }
    });
    byId("copyBtn").addEventListener("click", copyTranslation);
    byId("loadTableBtnTop").addEventListener("click", loadBatchTable);
    byId("loadTableBtnInline").addEventListener("click", loadBatchTable);
    byId("applyBatchTemplateBtn").addEventListener("click", applyBatchTemplate);
    byId("openWorkspaceConfigBtn").addEventListener("click", () => setWorkspaceConfigModal(true));
    byId("openWorkspaceConfigInlineBtn").addEventListener("click", () => setWorkspaceConfigModal(true));
    byId("closeWorkspaceConfigBtn").addEventListener("click", () => setWorkspaceConfigModal(false));
    byId("workspaceConfigModal").addEventListener("click", (event) => {
      if (event.target === byId("workspaceConfigModal")) {
        setWorkspaceConfigModal(false);
      }
    });
    byId("batch_message").addEventListener("input", resizeBatchMessageArea);
    byId("recordFilterFalse").addEventListener("click", () => {
      batchRecordFilter = "false";
      if (batchTablePayload) renderBatchTable(batchTablePayload);
    });
    byId("recordFilterTrue").addEventListener("click", () => {
      batchRecordFilter = "true";
      if (batchTablePayload) renderBatchTable(batchTablePayload);
    });
    byId("current_user").addEventListener("change", () => {
      syncRequesterWithCurrentUser(true);
      saveConfig();
    });
    Array.from(document.querySelectorAll(".workspace-tab")).forEach((el) => {
      el.addEventListener("click", () => {
        const target = el.dataset.panel;
        Array.from(document.querySelectorAll(".workspace-tab")).forEach((tab) => {
          tab.classList.toggle("active", tab === el);
        });
        Array.from(document.querySelectorAll(".workspace-panel")).forEach((panel) => {
          panel.classList.toggle("active", panel.id === target);
        });
      });
    });
    window.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        setWorkspaceConfigModal(false);
      }
    });
    sendHeartbeat();
    setInterval(sendHeartbeat, 3000);
    window.addEventListener("focus", sendHeartbeat);
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) {
        sendHeartbeat();
      }
    });
    loadState();
  </script>
</body>
</html>
"""


def load_saved_form() -> Dict[str, Any]:
    if not STATE_FILE.exists():
        return dict(DEFAULT_FORM)
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return dict(DEFAULT_FORM)
    form = dict(DEFAULT_FORM)
    if isinstance(data, dict):
        for key in form:
            if key in data:
                form[key] = data[key]
    return form

def persist_form(form: Dict[str, Any]) -> None:
    payload = dict(DEFAULT_FORM)
    for key in payload:
        if key in form:
            payload[key] = form[key]
    if not payload.get("save_crowdin_token"):
        payload["crowdin_token"] = ""
    if not payload.get("save_apitable_api_key"):
        payload["apitable_api_key"] = ""
    existing: Dict[str, Any] = {}
    if STATE_FILE.exists():
        try:
            existing = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if not isinstance(existing, dict):
                existing = {}
        except Exception:
            existing = {}
    existing.update(payload)
    STATE_FILE.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
    STATE["form"] = payload


def unwrap_resource(item: Dict[str, Any]) -> Dict[str, Any]:
    if not isinstance(item, dict):
        return {}
    data = item.get("data")
    if isinstance(data, dict):
        merged = dict(data)
        attrs = data.get("attributes")
        if isinstance(attrs, dict):
            merged.update(attrs)
        return merged
    attrs = item.get("attributes")
    if isinstance(attrs, dict):
        merged = dict(item)
        merged.update(attrs)
        return merged
    return item


def get_resource_id(item: Dict[str, Any]) -> str:
    value = unwrap_resource(item).get("id")
    return "" if value in (None, "") else str(value)


def get_resource_name(item: Dict[str, Any]) -> str:
    resource = unwrap_resource(item)
    for key in ("name", "title"):
        value = resource.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def get_resource_path(item: Dict[str, Any]) -> str:
    resource = unwrap_resource(item)
    for key in ("path", "filePath", "fullPath"):
        value = resource.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def ssl_context() -> ssl.SSLContext:
    context = ssl.create_default_context()
    if certifi is not None:
        context.load_verify_locations(cafile=certifi.where())
    return context


def http_json(
    method: str,
    url: str,
    *,
    headers: Optional[Dict[str, str]] = None,
    body: Optional[Dict[str, Any]] = None,
    timeout: int = 60,
) -> Dict[str, Any]:
    payload = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=payload, method=method, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_context()) as resp:
            raw = resp.read()
            return {} if not raw else json.loads(raw.decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} {exc.reason}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"网络请求失败: {exc}") from exc


def http_bytes(
    url: str,
    *,
    headers: Optional[Dict[str, str]] = None,
    timeout: int = 120,
) -> bytes:
    req = urllib.request.Request(url, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_context()) as resp:
            return resp.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code} {exc.reason}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"网络请求失败: {exc}") from exc


def crowdin_headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def crowdin_api_get_all(token: str, path: str, *, params: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    offset = 0
    while True:
        query = dict(params or {})
        query["limit"] = 500
        query["offset"] = offset
        url = f"{BASE_API_URL}{path}?{urllib.parse.urlencode(query)}"
        payload = http_json("GET", url, headers=crowdin_headers(token))
        page_items = payload.get("data") or []
        if not isinstance(page_items, list):
            break
        items.extend([unwrap_resource(item) for item in page_items if isinstance(item, dict)])
        if len(page_items) < 500:
            break
        offset += 500
    return items


def find_crowdin_project(token: str) -> Dict[str, Any]:
    for project in crowdin_api_get_all(token, "/projects"):
        identifier = str(project.get("identifier") or "").strip()
        name = str(project.get("name") or "").strip()
        if identifier == CROWDIN_PROJECT_SLUG or name == CROWDIN_PROJECT_SLUG:
            return project
    raise RuntimeError(f"未找到 Crowdin 项目: {CROWDIN_PROJECT_SLUG}")


def score_file_match(file_info: Dict[str, Any], file_name: str, extra_keyword: str) -> Optional[int]:
    path_value = get_resource_path(file_info).replace("\\", "/")
    name_value = get_resource_name(file_info)
    file_query = file_name.strip().lower()
    extra_query = extra_keyword.strip().lower()
    if not file_query:
        return None
    basename = path_value.rsplit("/", 1)[-1].lower()
    stem = basename.rsplit(".", 1)[0]
    haystack = f"{path_value} {name_value}".lower()
    if file_query not in haystack:
        return None
    if extra_query and extra_query not in haystack:
        return None
    if basename == file_query:
        return 400
    if stem == file_query:
        return 350
    if name_value.lower() == file_query:
        return 320
    if basename.startswith(file_query):
        return 260
    if f"/{file_query}" in path_value.lower():
        return 220
    return 180


def pick_crowdin_file(
    files: List[Dict[str, Any]],
    *,
    file_name: str,
    crowdin_folder: str,
    extra_path_keyword: str,
    extra_keyword: str,
) -> Dict[str, Any]:
    folder_needles = [crowdin_folder.strip(), extra_path_keyword.strip()]
    candidates: List[Tuple[int, Dict[str, Any]]] = []
    for file_info in files:
        path_value = get_resource_path(file_info).replace("\\", "/")
        lowered_path = path_value.lower()
        if any(needle and needle.lower() not in lowered_path for needle in folder_needles):
            continue
        score = score_file_match(file_info, file_name, extra_keyword)
        if score is None:
            continue
        candidates.append((score, file_info))
    if not candidates:
        raise LookupError("未搜到对应文件")
    candidates.sort(key=lambda item: (-item[0], get_resource_path(item[1]), get_resource_name(item[1])))
    return candidates[0][1]


def export_crowdin_translation(token: str, project_id: str, file_id: str) -> Dict[str, Any]:
    url = f"{BASE_API_URL}/projects/{project_id}/translations/builds/files/{file_id}"
    return http_json(
        "POST",
        url,
        headers=crowdin_headers(token),
        body={"targetLanguageId": CROWDIN_TARGET_LANGUAGE},
    )


def decode_translation_bytes(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "latin-1"):
        try:
            text = data.decode(encoding)
            if text:
                return text
        except Exception:
            continue
    return data.decode("utf-8", errors="replace")


def _strip_inner_xml(fragment: str) -> str:
    """Remove all XML tags from a fragment and return cleaned text."""
    fragment = re.sub(r"</?(?:\w+:)?(?:t|r|si|a|span|div|body|txBody|sheetData|row|c|is)[^>]*>", " ", fragment)
    fragment = re.sub(r"<[^>]+>", " ", fragment)
    fragment = (
        fragment.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )
    return " ".join(fragment.split()).strip()


def strip_xml_text(text: str) -> str:
    text = re.sub(r"<\?.*?\?>", " ", text, flags=re.S)

    # 尝试按段落标签 <w:p>/<p> 拆分，保留空段落产生的空行
    para_pattern = re.compile(r"<(?:\w+:)?p(?:\s[^>]*)?>.*?</(?:\w+:)?p\s*>|<(?:\w+:)?p(?:\s[^>]*)?/>", re.S)
    para_matches = list(para_pattern.finditer(text))

    if para_matches:
        # 先提取段落标签之前的文本（不在任何 <p> 里的内容）
        pre_text = _strip_inner_xml(text[:para_matches[0].start()])
        lines: list[str] = []
        if pre_text:
            lines.append(pre_text)
        for m in para_matches:
            inner = _strip_inner_xml(m.group())
            if inner:
                lines.append(inner)
            else:
                # 空段落 → 保留为空行，但避免开头和连续多个空行
                if lines and lines[-1] != "":
                    lines.append("")
        # 段落之后的尾部文本
        tail = _strip_inner_xml(text[para_matches[-1].end():])
        if tail:
            lines.append(tail)
        # 去掉末尾多余空行
        while lines and lines[-1] == "":
            lines.pop()
        return "\n".join(lines)

    # 非段落结构的 XML（如 xlsx sharedStrings 等），沿用原逻辑
    text = re.sub(r"</?(?:\w+:)?(?:t|r|si|a|span|div|body|txBody|sheetData|row|c|is)[^>]*>", "\n", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = (
        text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
    )
    lines2: list[str] = []
    for line in text.splitlines():
        cleaned = " ".join(line.split()).strip()
        if cleaned:
            lines2.append(cleaned)
    return "\n".join(lines2)


def archive_candidate_names(file_path: str) -> List[str]:
    lowered = file_path.lower()
    if lowered.endswith(".docx"):
        return ["word/document.xml"] + [f"word/{name}" for name in ("header1.xml", "footer1.xml", "footnotes.xml", "endnotes.xml")]
    if lowered.endswith(".pptx"):
        return [f"ppt/slides/slide{i}.xml" for i in range(1, 200)]
    if lowered.endswith(".xlsx"):
        names = ["xl/sharedStrings.xml"]
        names.extend([f"xl/worksheets/sheet{i}.xml" for i in range(1, 200)])
        return names
    return []


def extract_text_from_zip_bytes(data: bytes, file_path: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        all_names = set(zf.namelist())
        preferred = [name for name in archive_candidate_names(file_path) if name in all_names]
        if not preferred:
            preferred = [
                name for name in zf.namelist()
                if name.endswith((".txt", ".json", ".xml", ".html", ".htm"))
                and not name.startswith(("_rels/", "docProps/"))
                and not name.endswith((".rels",))
                and "[Content_Types]" not in name
            ]
        chunks: List[str] = []
        for name in preferred[:80]:
            try:
                raw = zf.read(name)
            except Exception:
                continue
            text = decode_translation_bytes(raw)
            if name.endswith(".xml") or "<" in text[:400]:
                text = strip_xml_text(text)
            else:
                text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
            if text:
                chunks.append(text)
        merged = "\n".join(chunk for chunk in chunks if chunk).strip()
        return merged


def looks_like_binary_garbage(text: str) -> bool:
    if not text:
        return False
    sample = text[:400]
    weird = 0
    for ch in sample:
        if ch == "\ufffd":
            weird += 3
        elif ord(ch) < 32 and ch not in "\n\r\t":
            weird += 2
    return weird > max(12, len(sample) // 12)


def extract_translation_text(data: bytes, file_path: str) -> str:
    if zipfile.is_zipfile(io.BytesIO(data)):
        text = extract_text_from_zip_bytes(data, file_path)
        if not text:
            raise RuntimeError("该文件导出为压缩包，但未提取到可读翻译内容，可能当前还没有翻译")
        return text
    text = decode_translation_bytes(data).strip()
    if not text:
        raise RuntimeError("导出的翻译内容为空，可能当前还没有翻译")
    if looks_like_binary_garbage(text):
        raise RuntimeError("导出的内容不是可直接写入的文本，可能是二进制文件或当前没有可读翻译")
    return text


def parse_apitable_workbench(url: str) -> Tuple[str, str, str]:
    parsed = urllib.parse.urlparse(url)
    parts = [item for item in parsed.path.split("/") if item]
    if len(parts) < 3 or parts[0] != "workbench":
        raise RuntimeError("APITable workbench 地址格式不正确")
    base_url = f"{parsed.scheme}://{parsed.netloc}/fusion/v1"
    return base_url, parts[1], parts[2]


def get_apitable_base_url() -> str:
    return f"{APITABLE_BASE_ORIGIN}/fusion/v1"


def build_multipart_body(fields: Dict[str, str], filename: str, file_bytes: bytes, content_type: str) -> Tuple[bytes, str]:
    boundary = f"----CodexBoundary{uuid.uuid4().hex}"
    buffer = io.BytesIO()
    for key, value in fields.items():
        buffer.write(f"--{boundary}\r\n".encode("utf-8"))
        buffer.write(f'Content-Disposition: form-data; name="{key}"\r\n\r\n'.encode("utf-8"))
        buffer.write(str(value).encode("utf-8"))
        buffer.write(b"\r\n")
    buffer.write(f"--{boundary}\r\n".encode("utf-8"))
    buffer.write(
        (
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f"Content-Type: {content_type}\r\n\r\n"
        ).encode("utf-8")
    )
    buffer.write(file_bytes)
    buffer.write(b"\r\n")
    buffer.write(f"--{boundary}--\r\n".encode("utf-8"))
    return buffer.getvalue(), boundary


def call_aitable(method: str, url: str, api_key: str, body: Optional[Dict[str, Any]] = None) -> Tuple[int, str]:
    headers = {"Authorization": f"Bearer {api_key}"}
    payload = None
    if body is not None:
        headers["Content-Type"] = "application/json"
        payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=payload, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=60, context=ssl_context()) as resp:
            return resp.status, resp.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        return exc.code, exc.read().decode("utf-8", errors="replace")
    except urllib.error.URLError as exc:
        return 500, f"网络错误: {exc}"


def upload_aitable_attachment(
    api_key: str,
    base_origin: str,
    datasheet_id: str,
    filename: str,
    file_bytes: bytes,
) -> List[Dict[str, Any]]:
    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    body, boundary = build_multipart_body(
        {},
        filename,
        file_bytes,
        content_type,
    )
    url = f"{base_origin}/fusion/v1/datasheets/{datasheet_id}/attachments"
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    }
    req = urllib.request.Request(url, data=body, method="POST", headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=120, context=ssl_context()) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"上传 APITable 附件失败 HTTP {exc.code}: {detail[:400]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"上传 APITable 附件失败: {exc}") from exc
    data = payload.get("data")
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if isinstance(data, dict):
        files = data.get("files") or data.get("attachments") or data.get("items")
        if isinstance(files, list):
            return [item for item in files if isinstance(item, dict)]
        return [data]
    raise RuntimeError(f"上传 APITable 附件成功，但返回结构无法识别: {json.dumps(payload, ensure_ascii=False)[:400]}")


def fetch_aitable_fields(api_key: str, base_url: str, datasheet_id: str, view_id: str) -> List[Dict[str, Any]]:
    query = urllib.parse.urlencode({"viewId": view_id})
    url = f"{base_url}/datasheets/{datasheet_id}/fields?{query}"
    status, text = call_aitable("GET", url, api_key)
    if status >= 400:
        raise RuntimeError(f"读取 APITable 字段失败 HTTP {status}: {text[:300]}")
    data = json.loads(text)
    return (data.get("data") or {}).get("fields") or []


def fetch_aitable_records(api_key: str, base_url: str, datasheet_id: str, view_id: str) -> List[Dict[str, Any]]:
    page = 1
    records: List[Dict[str, Any]] = []
    while True:
        query = urllib.parse.urlencode({"viewId": view_id, "pageSize": 1000, "pageNum": page})
        url = f"{base_url}/datasheets/{datasheet_id}/records?{query}"
        status, text = call_aitable("GET", url, api_key)
        if status >= 400:
            raise RuntimeError(f"读取 APITable 记录失败 HTTP {status}: {text[:300]}")
        data = json.loads(text)
        page_items = ((data.get("data") or {}).get("records")) or []
        if not isinstance(page_items, list):
            break
        records.extend(page_items)
        if len(page_items) < 1000:
            break
        page += 1
    return records


def field_by_name(fields: List[Dict[str, Any]], name: str) -> Optional[Dict[str, Any]]:
    for field in fields:
        if str(field.get("name") or "") == name:
            return field
    return None


def serialize_option_value(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def parse_field_options(field: Optional[Dict[str, Any]], records: List[Dict[str, Any]], field_name: str) -> List[Dict[str, Any]]:
    options: List[Dict[str, Any]] = []
    seen = set()
    has_schema_options = False
    if field:
        raw_options = ((field.get("property") or {}).get("options")) or ((field.get("properties") or {}).get("options")) or []
        if isinstance(raw_options, list):
            has_schema_options = bool(raw_options)
            for item in raw_options:
                if not isinstance(item, dict):
                    continue
                raw_value = item.get("id")
                if raw_value in (None, ""):
                    raw_value = item.get("name")
                label = item.get("name")
                if label in (None, ""):
                    label = raw_value
                if raw_value in (None, ""):
                    continue
                key = serialize_option_value(raw_value)
                if key in seen:
                    continue
                seen.add(key)
                options.append({"label": str(label), "value": raw_value})
    for record in records:
        raw_value = (record.get("fields") or {}).get(field_name)
        if raw_value in (None, "", []):
            continue
        if has_schema_options and isinstance(raw_value, str) and raw_value.lower().startswith("opt"):
            continue
        key = serialize_option_value(raw_value)
        if key in seen:
            continue
        seen.add(key)
        options.append({"label": normalize_cell_text(raw_value), "value": raw_value})
    return options


def resolve_option_value(
    fields: List[Dict[str, Any]],
    records: List[Dict[str, Any]],
    field_name: str,
    template_value: Any,
) -> Tuple[Any, Dict[str, Any]]:
    debug: Dict[str, Any] = {
        "field_name": field_name,
        "template_value": template_value,
        "matched_label": "",
        "matched_source": "",
        "matched_value": None,
        "option_labels": [],
    }
    if template_value in (None, "", []):
        return template_value, debug
    text_value = normalize_cell_text(template_value).strip()
    if not text_value:
        return template_value, debug
    options = parse_field_options(field_by_name(fields, field_name), records, field_name)
    debug["option_labels"] = [str(item.get("label") or "") for item in options]
    for item in options:
        if str(item.get("label") or "").strip() == text_value:
            debug["matched_label"] = text_value
            debug["matched_source"] = "options"
            debug["matched_value"] = item.get("value")
            return item.get("value"), debug
    for record in records:
        raw_value = (record.get("fields") or {}).get(field_name)
        if normalize_cell_text(raw_value).strip() == text_value:
            debug["matched_label"] = text_value
            debug["matched_source"] = "records"
            debug["matched_value"] = raw_value
            return raw_value, debug
    debug["matched_label"] = text_value
    debug["matched_source"] = "fallback"
    debug["matched_value"] = template_value
    return template_value, debug


def build_table_payload(
    fields: List[Dict[str, Any]],
    records: List[Dict[str, Any]],
    reference_records: Optional[List[Dict[str, Any]]] = None,
) -> Dict[str, Any]:
    table_columns = [
        SYNC_FIELD_SEARCH_KEY,
        READ_FIELD_REQUESTER,
        READ_FIELD_DEADLINE,
        READ_FIELD_PRIORITY,
        READ_FIELD_CROWDIN_FLAG,
        READ_FIELD_TASK_LINK,
        BATCH_FIELD_ASSIGN_LOCALIZER,
        BATCH_FIELD_REQUESTER,
        BATCH_FIELD_MESSAGE,
        BATCH_FIELD_CHECK_LEAD,
        BATCH_FIELD_SEND_FLAG,
    ]
    reference_map: Dict[str, Dict[str, Any]] = {}
    for item in reference_records or []:
        record_id = str(item.get("recordId") or "")
        if record_id:
            reference_map[record_id] = item
    rows = []
    for record in records:
        field_values = record.get("fields") or {}
        reference_values = (reference_map.get(str(record.get("recordId") or "")) or {}).get("fields") or {}
        def pick_value(field_name: str) -> Any:
            if field_name in field_values and field_values.get(field_name) not in (None, "", []):
                return field_values.get(field_name)
            return reference_values.get(field_name)
        rows.append({
            "recordId": record.get("recordId") or "",
            "cells": {name: normalize_cell_text(pick_value(name)) for name in table_columns},
            "rawFields": {name: pick_value(name) for name in table_columns},
            "summary": {
                "title": format_display_value(SYNC_FIELD_SEARCH_KEY, pick_value(SYNC_FIELD_SEARCH_KEY)),
                "requester": format_display_value(READ_FIELD_REQUESTER, pick_value(READ_FIELD_REQUESTER)),
                "deadline": format_display_value(READ_FIELD_DEADLINE, pick_value(READ_FIELD_DEADLINE)),
                "priority": format_display_value(READ_FIELD_PRIORITY, pick_value(READ_FIELD_PRIORITY)),
                "crowdin_flag": format_display_value(READ_FIELD_CROWDIN_FLAG, pick_value(READ_FIELD_CROWDIN_FLAG)),
                "task_link": format_display_value(READ_FIELD_TASK_LINK, pick_value(READ_FIELD_TASK_LINK)),
                "localizer": format_display_value(BATCH_FIELD_ASSIGN_LOCALIZER, pick_value(BATCH_FIELD_ASSIGN_LOCALIZER)),
                "owner": format_display_value(BATCH_FIELD_REQUESTER, pick_value(BATCH_FIELD_REQUESTER)),
                "message": format_display_value(BATCH_FIELD_MESSAGE, pick_value(BATCH_FIELD_MESSAGE)),
                "check_lead": format_display_value(BATCH_FIELD_CHECK_LEAD, pick_value(BATCH_FIELD_CHECK_LEAD)),
                "send_flag": format_display_value(BATCH_FIELD_SEND_FLAG, pick_value(BATCH_FIELD_SEND_FLAG)),
            },
        })
    meta = {}
    for field_name in table_columns:
        field = field_by_name(fields, field_name)
        meta[field_name] = {
            "type": str((field or {}).get("type") or ""),
            "options": parse_field_options(field, records, field_name),
        }
    return {
        "columns": table_columns,
        "rows": rows,
        "meta": meta,
    }


def normalize_record_fields_for_write(
    fields: List[Dict[str, Any]],
    records: List[Dict[str, Any]],
    values: Dict[str, Any],
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    type_map = {str(field.get("name") or ""): str(field.get("type") or "") for field in fields}
    debug_info: Dict[str, Any] = {"field_debug": {}}
    normalized: Dict[str, Any] = {}

    for field_name, raw_value in values.items():
        field_type = type_map.get(field_name, "")
        debug_info["field_debug"].setdefault(field_name, {})
        debug_info["field_debug"][field_name]["field_type"] = field_type
        if field_type == "Checkbox":
            normalized[field_name] = bool(raw_value)
            continue
        if field_type in ("Number", "Currency", "Percent", "Rating", "AutoNumber"):
            try:
                if raw_value in (None, ""):
                    normalized[field_name] = None
                else:
                    normalized[field_name] = float(str(raw_value))
            except Exception:
                normalized[field_name] = raw_value
            continue
        if field_type == "SingleSelect":
            resolved, field_debug = resolve_option_value(fields, records, field_name, raw_value)
            debug_info["field_debug"][field_name] = field_debug
            debug_info["field_debug"][field_name]["field_type"] = field_type
            normalized[field_name] = str(field_debug.get("matched_label") or normalize_cell_text(raw_value) or normalize_cell_text(resolved))
            continue
        if field_type == "MultiSelect":
            resolved, field_debug = resolve_option_value(fields, records, field_name, raw_value)
            debug_info["field_debug"][field_name] = field_debug
            debug_info["field_debug"][field_name]["field_type"] = field_type
            labels: List[str] = []
            if isinstance(raw_value, list):
                labels = [normalize_cell_text(item).strip() for item in raw_value if normalize_cell_text(item).strip()]
            else:
                label = str(field_debug.get("matched_label") or normalize_cell_text(raw_value) or normalize_cell_text(resolved)).strip()
                if label:
                    labels = [label]
            normalized[field_name] = labels
            continue
        if field_type in ("Member", "CreatedBy", "LastModifiedBy"):
            resolved, field_debug = resolve_option_value(fields, records, field_name, raw_value)
            debug_info["field_debug"][field_name] = field_debug
            debug_info["field_debug"][field_name]["field_type"] = field_type
            normalized[field_name] = resolved
            continue
        if isinstance(raw_value, (dict, list)):
            normalized[field_name] = normalize_cell_text(raw_value)
            continue
        normalized[field_name] = raw_value

    return normalized, debug_info


def apply_batch_template_to_record(
    api_key: str,
    base_url: str,
    datasheet_id: str,
    view_id: str,
    record_id: str,
    fields: List[Dict[str, Any]],
    records: List[Dict[str, Any]],
    template: Dict[str, Any],
) -> Dict[str, Any]:
    type_map = {str(field.get("name") or ""): str(field.get("type") or "") for field in fields}
    debug_info: Dict[str, Any] = {"field_debug": {}, "retry_payloads": []}

    def candidate_values(field_name: str, display_value: Any, resolved_value: Any) -> List[Any]:
        field_type = type_map.get(field_name, "")
        candidates: List[Any] = []

        def push(v: Any) -> None:
            if v in (None, ""):
                return
            for existing in candidates:
                if existing == v:
                    return
            candidates.append(v)

        push(resolved_value)
        push(display_value)
        if field_type == "Member":
            if isinstance(resolved_value, dict):
                push([resolved_value])
                if resolved_value.get("id"):
                    push({"id": resolved_value.get("id"), "type": "Member"})
                    push([{"id": resolved_value.get("id"), "type": "Member"}])
                    push(str(resolved_value.get("id")))
            elif isinstance(display_value, str):
                push(display_value)
        if field_type == "SingleSelect":
            candidates = []
            if isinstance(display_value, str):
                candidates.append(display_value)
            if isinstance(resolved_value, dict):
                if resolved_value.get("name"):
                    candidates.append(resolved_value.get("name"))
                if resolved_value.get("id"):
                    candidates.append(resolved_value.get("id"))
            for item in candidates:
                push(item)
        return candidates

    def record_field_matches(field_name: str, expected_display: Any) -> bool:
        latest_records = fetch_aitable_records(api_key, base_url, datasheet_id, view_id)
        latest = None
        for item in latest_records:
            if str(item.get("recordId") or "") == record_id:
                latest = item
                break
        if not latest:
            return False
        actual = normalize_cell_text((latest.get("fields") or {}).get(field_name)).strip()
        expected = normalize_cell_text(expected_display).strip()
        return bool(expected) and actual == expected

    immediate_fields, normalize_debug = normalize_record_fields_for_write(
        fields,
        records,
        {
            BATCH_FIELD_ASSIGN_LOCALIZER: template.get("assign_localizer"),
            BATCH_FIELD_REQUESTER: template.get("requester"),
            BATCH_FIELD_MESSAGE: template.get("message", ""),
            BATCH_FIELD_CHECK_LEAD: template.get("check_lead", "30"),
        },
    )
    debug_info["field_debug"].update(normalize_debug.get("field_debug") or {})
    update_record_comment(
        api_key,
        base_url,
        datasheet_id,
        view_id,
        record_id,
        "__batch__",
        immediate_fields,
    )

    for field_name in (BATCH_FIELD_ASSIGN_LOCALIZER, BATCH_FIELD_REQUESTER):
        display_value = template.get("assign_localizer") if field_name == BATCH_FIELD_ASSIGN_LOCALIZER else template.get("requester")
        if record_field_matches(field_name, display_value):
            continue
        resolved_value = immediate_fields.get(field_name)
        for candidate in candidate_values(field_name, display_value, resolved_value):
            debug_info["retry_payloads"].append({"field": field_name, "candidate": candidate})
            update_record_comment(
                api_key,
                base_url,
                datasheet_id,
                view_id,
                record_id,
                "__batch__",
                {field_name: candidate},
            )
            if record_field_matches(field_name, display_value):
                break

    time.sleep(3)
    update_record_comment(
        api_key,
        base_url,
        datasheet_id,
        view_id,
        record_id,
        "__batch__",
        normalize_record_fields_for_write(
            fields,
            records,
            {BATCH_FIELD_SEND_FLAG: template.get("send_flag", False)},
        )[0],
    )
    return debug_info


def write_fields_with_retry(
    api_key: str,
    base_url: str,
    datasheet_id: str,
    view_id: str,
    record_id: str,
    fields: List[Dict[str, Any]],
    records: List[Dict[str, Any]],
    input_values: Dict[str, Any],
) -> Dict[str, Any]:
    type_map = {str(field.get("name") or ""): str(field.get("type") or "") for field in fields}
    normalized_fields, normalize_debug = normalize_record_fields_for_write(fields, records, input_values)
    debug_info: Dict[str, Any] = {
        "field_debug": normalize_debug.get("field_debug") or {},
        "retry_payloads": [],
        "final_payload": dict(normalized_fields),
    }

    def field_matches(field_name: str, expected_display: Any) -> bool:
        latest_records = fetch_aitable_records(api_key, base_url, datasheet_id, view_id)
        latest = None
        for item in latest_records:
            if str(item.get("recordId") or "") == record_id:
                latest = item
                break
        if not latest:
            return False
        actual = normalize_cell_text((latest.get("fields") or {}).get(field_name)).strip()
        expected = normalize_cell_text(expected_display).strip()
        if expected.lower() in ("true", "false"):
            return actual.lower() == expected.lower()
        return bool(expected) and actual == expected

    def candidate_values(field_name: str, display_value: Any, resolved_value: Any) -> List[Any]:
        field_type = type_map.get(field_name, "")
        candidates: List[Any] = []

        def push(v: Any) -> None:
            if v in (None, ""):
                return
            for existing in candidates:
                if existing == v:
                    return
            candidates.append(v)

        push(resolved_value)
        push(display_value)
        if field_type == "SingleSelect":
            if isinstance(resolved_value, dict):
                push(resolved_value.get("name"))
                push(resolved_value.get("id"))
        if field_type == "MultiSelect":
            if isinstance(display_value, list):
                push(display_value)
            elif isinstance(display_value, str):
                push([display_value])
                push(display_value)
            if isinstance(resolved_value, list):
                push(resolved_value)
            elif isinstance(resolved_value, str):
                push([resolved_value])
        if field_type in ("Member", "CreatedBy", "LastModifiedBy") and isinstance(resolved_value, dict):
            push([resolved_value])
            if resolved_value.get("id"):
                push({"id": resolved_value.get("id"), "type": "Member"})
                push([{"id": resolved_value.get("id"), "type": "Member"}])
                push(str(resolved_value.get("id")))
        return candidates

    if normalized_fields:
        update_record_comment(
            api_key,
            base_url,
            datasheet_id,
            view_id,
            record_id,
            "__batch__",
            normalized_fields,
        )

    for field_name, display_value in input_values.items():
        if not field_matches(field_name, display_value):
            resolved_value = normalized_fields.get(field_name)
            for candidate in candidate_values(field_name, display_value, resolved_value):
                debug_info["retry_payloads"].append({"field": field_name, "candidate": candidate})
                update_record_comment(
                    api_key,
                    base_url,
                    datasheet_id,
                    view_id,
                    record_id,
                    "__batch__",
                    {field_name: candidate},
                )
                if field_matches(field_name, display_value):
                    break

    verification: Dict[str, bool] = {}
    for field_name, display_value in input_values.items():
        verification[field_name] = field_matches(field_name, display_value)
    debug_info["verification"] = verification
    return debug_info


def normalize_cell_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        parts = [normalize_cell_text(item) for item in value]
        return " ".join([part for part in parts if part]).strip()
    if isinstance(value, dict):
        for key in ("name", "label", "title", "text"):
            if value.get(key) not in (None, ""):
                return normalize_cell_text(value.get(key))
        if value.get("type") == "Member" and value.get("id") not in (None, ""):
            return str(value.get("id"))
        if "text" in value:
            return normalize_cell_text(value.get("text"))
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def format_display_value(field_name: str, value: Any) -> str:
    text = normalize_cell_text(value).strip()
    if not text:
        return ""
    if field_name == READ_FIELD_DEADLINE and re.fullmatch(r"\d{12,14}", text):
        try:
            ts = int(text)
            if ts > 10_000_000_000:
                ts = ts / 1000
            dt = time.localtime(ts)
            return time.strftime("%Y/%m/%d %H:%M", dt)
        except Exception:
            return text
    return text


def find_target_record(records: List[Dict[str, Any]], search_column: str, keyword: str) -> Dict[str, Any]:
    search = keyword.strip().lower()
    if not search:
        raise RuntimeError("APITable 搜索关键词不能为空")
    matched: List[Dict[str, Any]] = []
    exact: List[Dict[str, Any]] = []
    for record in records:
        fields = record.get("fields") or {}
        value = normalize_cell_text(fields.get(search_column)).strip()
        lowered = value.lower()
        if not lowered:
            continue
        if search in lowered:
            matched.append(record)
        if lowered == search:
            exact.append(record)
    if len(exact) == 1:
        return exact[0]
    if len(matched) == 1:
        return matched[0]
    if not matched:
        raise LookupError(f"APITable 中未找到包含关键词“{keyword}”的记录")
    sample = []
    for record in matched[:5]:
        fields = record.get("fields") or {}
        sample.append(normalize_cell_text(fields.get(search_column))[:60])
    raise RuntimeError(
        f"APITable 匹配到 {len(matched)} 行，请缩小关键词。示例: {' | '.join(sample)}"
    )


def update_record_comment(
    api_key: str,
    base_url: str,
    datasheet_id: str,
    view_id: str,
    record_id: str,
    comment_column: str,
    comment_value: Any,
) -> None:
    query = urllib.parse.urlencode({"viewId": view_id})
    url = f"{base_url}/datasheets/{datasheet_id}/records?{query}"
    body = {
        "records": [
            {
                "recordId": record_id,
                "fields": comment_value if comment_column == "__batch__" else {comment_column: comment_value},
            }
        ],
        "fieldKey": "name",
    }
    status, text = call_aitable("PATCH", url, api_key, body)
    if status >= 400:
        raise RuntimeError(f"写入 APITable 失败 HTTP {status}: {text[:300]}")


def new_job(payload: Dict[str, Any]) -> Dict[str, Any]:
    job_id = uuid.uuid4().hex
    job = {
        "id": job_id,
        "status": "pending",
        "message": "任务已创建",
        "logs": [],
        "result": {},
        "error_code": "",
        "created_at": int(time.time() * 1000),
        "payload": payload,
    }
    STATE["jobs"][job_id] = job
    return job


def append_job_log(job: Dict[str, Any], text: str) -> None:
    job["logs"].append(text)
    job["message"] = text


def run_sync_job(job: Dict[str, Any]) -> None:
    payload = job["payload"]
    job["status"] = "running"
    try:
        crowdin_token = payload["crowdin_token"].strip()
        apitable_api_key = payload["apitable_api_key"].strip()
        current_user = payload.get("current_user", "").strip()
        sync_note = payload.get("sync_note", "")
        file_name = payload["file_name"].strip()
        crowdin_folder = payload["crowdin_folder"].strip() or "English Team"
        extra_path_keyword = payload["extra_path_keyword"].strip()
        extra_keyword = payload["extra_keyword"].strip()
        search_keyword = file_name
        search_column = (
            payload["apitable_search_column"].strip()
            or "翻译需求（当日日期+翻译需求名字+需求人）"
        )
        comment_column = (
            payload["apitable_comment_column"].strip()
            or "完成版附件"
        )

        append_job_log(job, "1/6 正在定位 Crowdin 项目…")
        project = find_crowdin_project(crowdin_token)
        project_id = get_resource_id(project)

        append_job_log(job, "2/6 正在读取 Crowdin 文件清单…")
        files = crowdin_api_get_all(crowdin_token, f"/projects/{project_id}/files", params={"recursion": "true"})

        append_job_log(job, "3/6 正在匹配目标文件…")
        try:
            matched_file = pick_crowdin_file(
                files,
                file_name=file_name,
                crowdin_folder=crowdin_folder,
                extra_path_keyword=extra_path_keyword,
                extra_keyword=extra_keyword,
            )
        except LookupError as exc:
            job["error_code"] = "crowdin_not_found"
            raise RuntimeError(str(exc)) from exc
        matched_file_path = get_resource_path(matched_file) or get_resource_name(matched_file)
        append_job_log(job, f"已匹配文件: {matched_file_path}")

        append_job_log(job, "4/6 正在导出 en-US 翻译…")
        export_result = export_crowdin_translation(crowdin_token, project_id, get_resource_id(matched_file))
        data = export_result.get("data") or {}
        download_url = str(data.get("url") or "").strip()
        if not download_url:
            raise RuntimeError(f"Crowdin 未返回下载链接: {json.dumps(export_result, ensure_ascii=False)[:500]}")
        raw_translation = http_bytes(download_url)
        attachment_name = os.path.basename(matched_file_path.replace("\\", "/")) or f"{file_name}.bin"
        if zipfile.is_zipfile(io.BytesIO(raw_translation)):
            append_job_log(job, "Crowdin 返回的是压缩包类型，正在提取可读文本…")
        else:
            append_job_log(job, "Crowdin 返回的是文本类型，正在整理内容…")
        try:
            translation_text = extract_translation_text(raw_translation, matched_file_path).strip()
        except Exception as exc:
            translation_text = ""
            append_job_log(job, f"未提取到可读预览，将继续上传原文件: {exc}")

        append_job_log(job, "5/6 正在读取 APITable 视图数据…")
        base_url = get_apitable_base_url()
        datasheet_id = APITABLE_DATASHEET_ID
        view_id = APITABLE_WRITE_VIEW_ID
        base_origin = base_url.rsplit("/fusion/v1", 1)[0]
        fields = fetch_aitable_fields(apitable_api_key, base_url, datasheet_id, view_id)
        field_names = {str(field.get("name") or "") for field in fields}
        if search_column not in field_names:
            raise RuntimeError(f"APITable 中未找到搜索列: {search_column}")
        if comment_column not in field_names:
            raise RuntimeError(f"APITable 中未找到写入列: {comment_column}")
        records = fetch_aitable_records(apitable_api_key, base_url, datasheet_id, view_id)
        record = find_target_record(records, search_column, search_keyword)
        record_id = str(record.get("recordId") or "")
        if not record_id:
            raise RuntimeError("匹配记录缺少 recordId")

        checker_field_payload = {}
        if SYNC_FIELD_CHECKER in field_names and current_user:
            checker_field_payload[SYNC_FIELD_CHECKER] = current_user
        if checker_field_payload:
            append_job_log(job, "6/6 先写入检查者…")
            checker_write_debug = write_fields_with_retry(
                apitable_api_key,
                base_url,
                datasheet_id,
                view_id,
                record_id,
                fields,
                records,
                checker_field_payload,
            )
            append_job_log(job, f"已写入检查者: {json.dumps(checker_write_debug.get('final_payload') or {}, ensure_ascii=False)}")
            if checker_write_debug.get("field_debug"):
                append_job_log(job, f"检查者字段映射: {json.dumps(checker_write_debug['field_debug'], ensure_ascii=False)}")
            if checker_write_debug.get("retry_payloads"):
                append_job_log(job, f"检查者重试: {json.dumps(checker_write_debug['retry_payloads'], ensure_ascii=False)}")
            if checker_write_debug.get("verification"):
                append_job_log(job, f"检查者校验: {json.dumps(checker_write_debug['verification'], ensure_ascii=False)}")
            time.sleep(1)

        append_job_log(job, "正在上传附件并写入 APITable…")
        existing_attachments = []
        existing_value = (record.get("fields") or {}).get(comment_column)
        if isinstance(existing_value, list):
            existing_attachments = [item for item in existing_value if isinstance(item, dict)]
        uploaded_attachments = upload_aitable_attachment(
            apitable_api_key,
            base_origin,
            datasheet_id,
            attachment_name,
            raw_translation,
        )
        if not uploaded_attachments:
            raise RuntimeError("APITable 附件上传后没有返回可写入的附件对象")
        update_record_comment(
            apitable_api_key,
            base_url,
            datasheet_id,
            view_id,
            record_id,
            comment_column,
            existing_attachments + uploaded_attachments,
        )
        extra_sync_fields = {}
        if SYNC_FIELD_CHECKED in field_names:
            extra_sync_fields[SYNC_FIELD_CHECKED] = True
        if SYNC_FIELD_NOTE in field_names and sync_note:
            extra_sync_fields[SYNC_FIELD_NOTE] = sync_note
        if extra_sync_fields:
            sync_field_types = {
                name: str((field_by_name(fields, name) or {}).get("type") or "")
                for name in extra_sync_fields
            }
            append_job_log(job, f"检查信息字段类型: {json.dumps(sync_field_types, ensure_ascii=False)}")
            sync_write_debug = write_fields_with_retry(
                apitable_api_key,
                base_url,
                datasheet_id,
                view_id,
                record_id,
                fields,
                records,
                extra_sync_fields,
            )
            append_job_log(job, f"已补写检查信息: {json.dumps(sync_write_debug.get('final_payload') or {}, ensure_ascii=False)}")
            if sync_write_debug.get("field_debug"):
                append_job_log(job, f"检查信息字段映射: {json.dumps(sync_write_debug['field_debug'], ensure_ascii=False)}")
            if sync_write_debug.get("retry_payloads"):
                append_job_log(job, f"检查信息重试: {json.dumps(sync_write_debug['retry_payloads'], ensure_ascii=False)}")
            if sync_write_debug.get("verification"):
                append_job_log(job, f"检查信息校验: {json.dumps(sync_write_debug['verification'], ensure_ascii=False)}")

        job["status"] = "success"
        job["message"] = "同步完成，完成版附件已更新"
        job["result"] = {
            "matched_file_path": matched_file_path,
            "record_id": record_id,
            "search_column": search_column,
            "comment_column": comment_column,
            "translation_text": translation_text or f"[附件已上传] {attachment_name}",
            "attachment_name": attachment_name,
        }
        append_job_log(job, "同步完成，完成版附件已更新。")
    except Exception as exc:
        job["status"] = "error"
        job["message"] = str(exc)
        append_job_log(job, f"[ERROR] {exc}")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def send_html(self, text: str, status: int = 200) -> None:
        body = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, data: Any, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_json(self) -> Dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0") or 0)
        raw = self.rfile.read(length).decode("utf-8", errors="replace")
        return json.loads(raw) if raw else {}

    def do_GET(self) -> None:
        if self.path in ("/", "/index.html"):
            self.send_html(INDEX_HTML)
            return
        if self.path == "/api/state":
            self.send_json({"form": STATE["form"]})
            return
        if self.path.startswith("/api/job/"):
            job_id = self.path.rsplit("/", 1)[-1]
            job = STATE["jobs"].get(job_id)
            if not job:
                self.send_json({"status": "error", "message": "任务不存在"}, 404)
                return
            self.send_json({
                "id": job["id"],
                "status": job["status"],
                "message": job["message"],
                "logs": job["logs"],
                "result": job["result"],
                "error_code": job.get("error_code", ""),
            })
            return
        self.send_error(404)

    def do_POST(self) -> None:
        if self.path == "/api/ping":
            mark_client_ping()
            self.send_json({"ok": True})
            return
        if self.path == "/api/shutdown":
            self.send_json({"ok": True})
            request_shutdown()
            return
        if self.path == "/api/save-config":
            try:
                payload = self.read_json()
                merged = dict(DEFAULT_FORM)
                merged.update(payload)
                persist_form(merged)
                self.send_json({"ok": True})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        if self.path == "/api/batch-table/load":
            try:
                payload = self.read_json()
                api_key = str(payload.get("api_key") or "").strip()
                if not api_key:
                    raise RuntimeError("APITable API Key 不能为空")
                base_url = get_apitable_base_url()
                datasheet_id = APITABLE_DATASHEET_ID
                fields = fetch_aitable_fields(api_key, base_url, datasheet_id, APITABLE_WRITE_VIEW_ID)
                records = fetch_aitable_records(api_key, base_url, datasheet_id, APITABLE_WRITE_VIEW_ID)
                reference_records = fetch_aitable_records(api_key, base_url, datasheet_id, APITABLE_READ_VIEW_ID)
                self.send_json({"ok": True, "table": build_table_payload(fields, records, reference_records)})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        if self.path == "/api/batch-apply":
            try:
                payload = self.read_json()
                api_key = str(payload.get("api_key") or "").strip()
                record_id = str(payload.get("record_id") or "").strip()
                template = payload.get("template") or {}
                if not api_key:
                    raise RuntimeError("APITable API Key 不能为空")
                if not record_id:
                    raise RuntimeError("请先选中一行")
                base_url = get_apitable_base_url()
                datasheet_id = APITABLE_DATASHEET_ID
                fields = fetch_aitable_fields(api_key, base_url, datasheet_id, APITABLE_WRITE_VIEW_ID)
                records = fetch_aitable_records(api_key, base_url, datasheet_id, APITABLE_WRITE_VIEW_ID)
                debug_info = apply_batch_template_to_record(api_key, base_url, datasheet_id, APITABLE_WRITE_VIEW_ID, record_id, fields, records, template)
                self.send_json({
                    "ok": True,
                    "message": "前四项已写入，勾选字段已在 3 秒后完成更新",
                    "debug": debug_info,
                })
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        if self.path == "/api/start-sync":
            try:
                payload = self.read_json()
                merged = dict(DEFAULT_FORM)
                merged.update(payload)
                merged["crowdin_folder"] = merged["crowdin_folder"].strip() or "English Team"
                if not merged["crowdin_token"].strip():
                    raise RuntimeError("Crowdin Token 不能为空")
                if not merged["apitable_api_key"].strip():
                    raise RuntimeError("APITable API Key 不能为空")
                if not merged["file_name"].strip():
                    raise RuntimeError("文件名或主关键词不能为空")
                persist_form(merged)
                job = new_job(merged)
                thread = threading.Thread(target=run_sync_job, args=(job,), daemon=True)
                thread.start()
                self.send_json({"ok": True, "job_id": job["id"]})
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, 400)
            return
        self.send_error(404)


def main() -> None:
    global SERVER_REF
    STATE["form"] = load_saved_form()
    server, selected_port = create_server_with_fallback(HOST, PORT)
    SERVER_REF = server
    url = f"http://{HOST}:{selected_port}"
    if not getattr(sys, "frozen", False):
        print(f"Crowdin → APITable 留言同步工具已启动: {url}")
        print(f"读取视图: {APITABLE_DATASHEET_ID}/{APITABLE_READ_VIEW_ID}")
        print(f"写入视图: {APITABLE_DATASHEET_ID}/{APITABLE_WRITE_VIEW_ID}")
        print(f"本地配置文件: {STATE_FILE}")
    threading.Thread(target=watchdog_auto_exit, daemon=True).start()
    threading.Timer(0.4, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        if not getattr(sys, "frozen", False):
            print("\n已退出")


if __name__ == "__main__":
    main()
