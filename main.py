#!/usr/bin/env python3
"""
Crowdin 翻译任务新建工具
填写表单 → 写入 APITable → (可选) 上传文件到 Crowdin English Team
"""
from __future__ import annotations

import json
import mimetypes
import os
import re
import ssl
import socket
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import webbrowser
import zipfile
from html.parser import HTMLParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── 常量 ──────────────────────────────────────────────────────────────────────
HOST = "127.0.0.1"
PORT = 8773
CLIENT_HEARTBEAT_TIMEOUT = 12
CLIENT_HEARTBEAT_INTERVAL = 3

CROWDIN_BASE_API    = "https://api.crowdin.com/api/v2"
CROWDIN_PROJECT_SLUG = "operational-localization"
CROWDIN_DEFAULT_LINK = "https://crowdin.com/editor/operational-localization/49563/zhcn-enus?view=comfortable&filter=basic&value=3"
CROWDIN_EDITOR_LOCALE = "zhcn-enus"
CROWDIN_EDITOR_QUERY = "view=side-by-side&filter=advanced&value=12&verbal_expression_scope=translations"
CROWDIN_UPLOAD_FOLDER = "English Team"

APITABLE_BASE         = "https://apitable.yottastudios.com/fusion/v1"
APITABLE_DATASHEET_ID = "dst54Y1Wzwdm5sDeQ7"
APITABLE_VIEW_ID      = "viw8QB941rgSg"
APITABLE_DEFAULT_KEY  = ""

# APITable 字段名
F_TITLE    = "翻译需求（当日日期+翻译需求名字+需求人）"
F_REQ      = "需求人"
F_DDL      = "需求ddl（尽量提前1-3小时）"
F_PRIORITY = "优先级"
F_SWITCH   = "是否用运营组专用crowdin"
F_LINK     = "填写实际任务链接，默认crowdin链接"
F_NOTE     = "备注(不放具体链接)"
F_UPLOADED = "是否上传到crowdin（决定信息是否发送）"

REQUESTERS = [
    "八幡","无心","石上","江阔","愿望","郭嘉","尤姆","新八",
    "不知火","卢仙","芝士","云逸","莉瑟","钟鱼","白叶",
    "仿生人","德古拉","未来","莱因","脆脆鲨","知南","心月狐",
]

if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).resolve().parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent

STATE_FILE = SCRIPT_DIR / "crowdin_tool_state.json"

DEFAULT_STATE: Dict[str, Any] = {
    "crowdin_token":        "",
    "save_crowdin_token":   True,
    "apitable_api_key":     APITABLE_DEFAULT_KEY,
    "save_apitable_key":    True,
    "current_user":         "",
    "crowdin_folder":       "English Team",
    "crowdin_language":     "en-US",
    "extra_path_keyword":   "",
    "extra_keyword":        "",
}

STATE: Dict[str, Any] = {}
JOBS:  Dict[str, Any] = {}
LAST_CLIENT_SEEN = time.time()

def touch_client() -> None:
    global LAST_CLIENT_SEEN
    LAST_CLIENT_SEEN = time.time()

def has_active_jobs() -> bool:
    return any(j.get("status") in {"pending", "running"} for j in JOBS.values())

def client_watchdog(server: ThreadingHTTPServer) -> None:
    while True:
        time.sleep(CLIENT_HEARTBEAT_INTERVAL)
        idle_for = time.time() - LAST_CLIENT_SEEN
        if idle_for > CLIENT_HEARTBEAT_TIMEOUT and not has_active_jobs():
            server.shutdown()
            return

# ── SSL ───────────────────────────────────────────────────────────────────────
def ssl_ctx() -> ssl.SSLContext:
    try:
        import certifi  # type: ignore
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode    = ssl.CERT_NONE
        return ctx

# ── 持久化 ────────────────────────────────────────────────────────────────────
def load_state() -> Dict[str, Any]:
    if STATE_FILE.exists():
        try:
            d = json.loads(STATE_FILE.read_text("utf-8"))
            m = dict(DEFAULT_STATE); m.update(d); return m
        except Exception:
            pass
    return dict(DEFAULT_STATE)

def save_state(data: Dict[str, Any]) -> None:
    s = dict(data)
    save_crowdin = bool(s.get("save_crowdin_token", s.get("save_crowdin", True)))
    save_apitable = bool(s.get("save_apitable_key", s.get("save_aitable", True)))
    s["save_crowdin_token"] = save_crowdin
    s["save_apitable_key"] = save_apitable
    s["save_crowdin"] = save_crowdin
    s["save_aitable"] = save_apitable
    if not save_crowdin:  s["crowdin_token"]    = ""
    if not save_apitable: s["apitable_api_key"] = ""
    STATE_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), "utf-8")

def choose_port(preferred: int = PORT) -> int:
    for candidate in (preferred, 0):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind((HOST, candidate))
            except OSError:
                continue
            return int(sock.getsockname()[1])
    raise RuntimeError("无法找到可用本地端口")

# ── HTTP 工具 ─────────────────────────────────────────────────────────────────
def http_json(method: str, url: str, headers: Optional[Dict]=None,
              body: Any=None, timeout: int=60) -> Any:
    payload = None
    hdrs = dict(headers or {})
    if body is not None:
        payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
        hdrs.setdefault("Content-Type","application/json")
    req = urllib.request.Request(url, data=payload, method=method, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx()) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} {url}\n{detail[:600]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"网络错误: {e}") from e

# ── APITable ──────────────────────────────────────────────────────────────────
def aitable_hdr(key: str) -> Dict[str,str]:
    return {"Authorization": f"Bearer {key}"}

def aitable_ping(key: str) -> bool:
    """静默验证 APITable 连接"""
    try:
        r = http_json("GET",
            f"{APITABLE_BASE}/datasheets/{APITABLE_DATASHEET_ID}/records"
            f"?viewId={APITABLE_VIEW_ID}&pageSize=1",
            headers=aitable_hdr(key), timeout=10)
        return r.get("success") is True
    except Exception:
        return False

def aitable_create_record(key: str, fields: Dict[str, Any]) -> str:
    """新建一条记录，返回 recordId"""
    r = http_json("POST",
        f"{APITABLE_BASE}/datasheets/{APITABLE_DATASHEET_ID}/records",
        headers=aitable_hdr(key),
        body={"records": [{"fields": fields}]})
    if not r.get("success"):
        raise RuntimeError(f"APITable 写入失败: {json.dumps(r, ensure_ascii=False)[:300]}")
    return r["data"]["records"][0]["recordId"]

def aitable_patch_record(key: str, record_id: str, fields: Dict[str, Any]) -> None:
    http_json("PATCH",
        f"{APITABLE_BASE}/datasheets/{APITABLE_DATASHEET_ID}/records",
        headers=aitable_hdr(key),
        body={"records": [{"recordId": record_id, "fields": fields}]})

# ── Crowdin ───────────────────────────────────────────────────────────────────
def crowdin_hdr(token: str) -> Dict[str,str]:
    return {"Authorization": f"Bearer {token}"}

def crowdin_storage_filename_header(filename: str) -> str:
    return urllib.parse.quote(filename or "file", safe="")

def crowdin_get_all(token: str, path: str, params: Optional[Dict]=None) -> List[Any]:
    results: List[Any] = []
    limit, offset = 500, 0
    while True:
        qp = dict(params or {}); qp.update({"limit": limit, "offset": offset})
        url = f"{CROWDIN_BASE_API}{path}?{urllib.parse.urlencode(qp)}"
        resp = http_json("GET", url, headers=crowdin_hdr(token))
        data = resp.get("data") or []
        if not isinstance(data, list): break
        for item in data:
            results.append(item.get("data") if isinstance(item,dict) and "data" in item else item)
        if len(data) < limit: break
        offset += limit
    return results

def crowdin_find_project(token: str) -> Dict[str, Any]:
    for p in crowdin_get_all(token, "/projects"):
        if CROWDIN_PROJECT_SLUG in str(p.get("identifier","")).lower() \
        or CROWDIN_PROJECT_SLUG in str(p.get("name","")).lower():
            return p
    raise RuntimeError(f"找不到 Crowdin 项目 '{CROWDIN_PROJECT_SLUG}'，请检查 Token")

def crowdin_find_folder(token: str, project_id: str, name: str) -> Optional[str]:
    for f in crowdin_get_all(token, f"/projects/{project_id}/directories", {"recursion":"true"}):
        if str(f.get("name","")).strip().lower() == name.strip().lower():
            return str(f.get("id",""))
    return None

def crowdin_upload_storage(token: str, filename: str, data: bytes) -> str:
    ct = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    url = f"{CROWDIN_BASE_API}/storages"
    hdrs = {
        **crowdin_hdr(token),
        "Crowdin-API-FileName": crowdin_storage_filename_header(filename),
        "Content-Type": ct,
    }
    req = urllib.request.Request(url, data=data, method="POST", headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=120, context=ssl_ctx()) as r:
            result = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Storage 上传失败 HTTP {e.code}: {e.read().decode()[:400]}") from e
    sid = str((result.get("data") or result).get("id",""))
    if not sid: raise RuntimeError(f"Crowdin Storage 未返回 id: {result}")
    return sid

def crowdin_create_file(token: str, project_id: str, storage_id: str,
                        filename: str, dir_id: Optional[str]) -> Dict[str, Any]:
    body: Dict[str, Any] = {"storageId": int(storage_id), "name": filename}
    if dir_id: body["directoryId"] = int(dir_id)
    r = http_json("POST", f"{CROWDIN_BASE_API}/projects/{project_id}/files",
                  headers=crowdin_hdr(token), body=body)
    return r.get("data") if isinstance(r.get("data"),dict) else r

# ── HTML → DOCX 转换 ────────────────────────────────────────────────────────
class _HtmlTextExtractor(HTMLParser):
    """简单 HTML → 文本段落提取器，保留换行和基本格式标记"""
    def __init__(self):
        super().__init__()
        self.paragraphs: List[str] = []
        self._buf: List[str] = []
        self._in_block = False

    def _flush(self):
        text = "".join(self._buf).strip()
        if text:
            self.paragraphs.append(text)
        self._buf = []

    def handle_starttag(self, tag, attrs):
        if tag in ("br",):
            self._flush()
        elif tag in ("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "tr"):
            self._flush()

    def handle_endtag(self, tag):
        if tag in ("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "tr"):
            self._flush()

    def handle_data(self, data):
        self._buf.append(data)

    def close(self):
        super().close()
        self._flush()


def html_to_docx_bytes(html_content: str) -> bytes:
    """将 HTML 富文本内容转换为 .docx 文件字节。
    使用纯标准库（zipfile）手动构建 OOXML 格式。
    """
    # 提取纯文本段落
    parser = _HtmlTextExtractor()
    parser.feed(html_content)
    parser.close()
    paragraphs = parser.paragraphs or [""]

    # 构建段落 XML
    para_xml_parts = []
    for p in paragraphs:
        # 转义 XML 特殊字符
        escaped = p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        para_xml_parts.append(
            f'<w:p><w:r><w:t xml:space="preserve">{escaped}</w:t></w:r></w:p>'
        )
    body_xml = "\n".join(para_xml_parts)

    # OOXML 模板
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {body_xml}
  </w:body>
</w:document>"""

    word_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"""

    import io as _io
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/_rels/document.xml.rels", word_rels)
    return buf.getvalue()


# ── 后台 Job ──────────────────────────────────────────────────────────────────
def jlog(job: Dict, msg: str) -> None:
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    job["log"].append(line)
    print(line)

def run_submit_job(job: Dict[str, Any]) -> None:
    job["status"] = "running"
    try:
        p              = job["payload"]
        api_key        = str(p.get("apitable_api_key","")).strip()
        crowdin_token  = str(p.get("crowdin_token","")).strip()
        use_crowdin    = p.get("use_crowdin", False)          # 是否用运营组专用 crowdin
        fields         = p.get("fields", {})                  # APITable 字段
        file_bytes: bytes = p.get("file_bytes") or b""
        filename       = str(p.get("filename","")).strip()
        rich_html      = str(p.get("rich_html","")).strip()

        # 如果有富文本内容且没有文件，自动生成 docx
        if rich_html and not file_bytes:
            jlog(job, "  检测到纯文本内容，正在生成 .docx…")
            file_bytes = html_to_docx_bytes(rich_html)
            if not filename.endswith(".docx"):
                filename = filename or "document.docx"
            jlog(job, f"  ✓ 已生成 {filename}（{len(file_bytes):,} bytes）")
        upload_folder  = str(p.get("upload_folder", CROWDIN_UPLOAD_FOLDER)).strip()

        # ── Step 1: 写入 APITable ──────────────────────────────────────────
        jlog(job, "1/3 正在写入 APITable…")
        record_id = aitable_create_record(api_key, fields)
        jlog(job, f"✓ 记录已创建（{record_id}）")

        if not use_crowdin:
            # 不用运营组 crowdin → 也勾选 uploaded，表示此条已完成提交流程
            jlog(job, "  （非运营组 Crowdin，无需上传附件）")
            aitable_patch_record(api_key, record_id, {F_UPLOADED: True})
            jlog(job, "✓ 已勾选「是否上传到crowdin」")
            jlog(job, "✅ 完成！")
            job["status"] = "success"
            job["result"] = {"record_id": record_id, "uploaded": True}
            return

        # ── Step 2: 上传到 Crowdin ─────────────────────────────────────────
        if not crowdin_token:
            raise RuntimeError("请填写 Crowdin API Token")
        if not file_bytes or not filename:
            raise RuntimeError("请选择要上传的文件")

        jlog(job, "2/3 正在定位 Crowdin 项目…")
        project    = crowdin_find_project(crowdin_token)
        project_id = str(project.get("id",""))
        jlog(job, f"✓ 项目：{project.get('name','')}（id={project_id}）")

        jlog(job, f"  上传文件到 Storage（{len(file_bytes):,} bytes）…")
        storage_id = crowdin_upload_storage(crowdin_token, filename, file_bytes)
        jlog(job, f"✓ Storage 上传成功（id={storage_id}）")

        dir_id = crowdin_find_folder(crowdin_token, project_id, upload_folder)
        if dir_id:
            jlog(job, f"  找到文件夹 [{upload_folder}]（id={dir_id}）")
        else:
            jlog(job, f"  ⚠ 未找到文件夹 [{upload_folder}]，上传到根目录")

        result_file = crowdin_create_file(crowdin_token, project_id, storage_id, filename, dir_id)
        new_fid     = str(result_file.get("id",""))
        file_path   = str(result_file.get("path") or result_file.get("name") or filename)
        file_url    = f"https://crowdin.com/editor/{CROWDIN_PROJECT_SLUG}/{urllib.parse.quote(new_fid)}/{CROWDIN_EDITOR_LOCALE}?{CROWDIN_EDITOR_QUERY}"
        jlog(job, f"✓ 文件已创建：{file_path}")

        # ── Step 3: 回写 APITable uploaded=True ────────────────────────────
        jlog(job, "3/3 回写 APITable 上传状态…")
        aitable_patch_record(api_key, record_id, {F_UPLOADED: True, F_LINK: file_url})
        jlog(job, "✓ 已勾选「是否上传到crowdin」并回写对应 Crowdin 链接")

        jlog(job, "✅ 全部完成！")
        job["status"] = "success"
        job["result"] = {
            "record_id":  record_id,
            "uploaded":   True,
            "filename":   filename,
            "file_path":  file_path,
            "file_url":   file_url,
        }

    except Exception as exc:
        job["status"] = "error"
        job["error"]  = str(exc)
        jlog(job, f"❌ 失败：{exc}")

# ── HTML ──────────────────────────────────────────────────────────────────────
def build_html() -> str:
    req_options = "\n".join(
        f'<option value="{n}">{n}</option>' for n in REQUESTERS
    )
    crowdin_url = CROWDIN_DEFAULT_LINK.replace("&", "\\x26")

    _JS = "// \u2500\u2500 state \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\nlet fileBytes = null, fileName = \"\", useCrowdin = null, jobId = \"\", poll = null;\nconst byId = id => document.getElementById(id);\nconst v    = id => (byId(id)||{}).value || \"\";\nconst esc  = s => String(s||\"\").replace(/&/g,\"&amp;\").replace(/</g,\"&lt;\").replace(/>/g,\"&gt;\");\n\nconst CROWDIN_DEFAULT = \"__CROWDIN_URL__\";\n\n// \u2500\u2500 init \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\n// \u2500\u2500 \u5173\u95ed\u6d4f\u89c8\u5668\u65f6\u81ea\u52a8\u5173\u95ed\u670d\u52a1 \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\nwindow.addEventListener(\"beforeunload\", () => {\n  navigator.sendBeacon(\"/api/shutdown\");\n});\n\nwindow.addEventListener(\"load\", async () => {\n  // \u81ea\u52a8\u586b\u4eca\u5929\u65e5\u671f\u5230\u6807\u9898\n  const today = new Date();\n  const pad = n => String(n).padStart(2,\"0\");\n  const ds = `${today.getFullYear()}${pad(today.getMonth()+1)}${pad(today.getDate())}`;\n  byId(\"f_title\").placeholder = `\u683c\u5f0f\uff1a${ds} \u7ffb\u8bd1\u9700\u6c42\u540d\u5b57 \u9700\u6c42\u4eba`;\n\n  try {\n    const cfg = await api(\"/api/state\");\n    applyConfig(cfg);\n    // \u9759\u9ed8\u68c0\u67e5\u8fde\u63a5\n    checkConn(cfg.apitable_api_key);\n  } catch(_) {}\n});\n\nfunction applyConfig(d) {\n  if (!d) return;\n  [\"current_user\",\"crowdin_token\",\"apitable_api_key\",\n   \"crowdin_folder\",\"extra_path_keyword\",\"extra_keyword\"].forEach(k => {\n    const el = byId(k); if (el && d[k]) el.value = d[k];\n  });\n  if (d.save_crowdin != null) byId(\"save_crowdin\").checked = !!d.save_crowdin;\n  if (d.save_aitable != null) byId(\"save_aitable\").checked = !!d.save_aitable;\n  // \u540c\u6b65 req \u4e0e current_user\n  if (d.current_user) byId(\"f_req\").value = d.current_user;\n  // \u66f4\u65b0\u6587\u4ef6\u5939 label\n  updateFolderLabel();\n};\n\nasync function checkConn(key) {\n  if (!key) { setConn(false); return; }\n  try {\n    const r = await api(\"/api/ping\", {apitable_api_key: key});\n    setConn(r.ok === true);\n  } catch(_) { setConn(false); }\n}\n\nfunction setConn(ok) {\n  byId(\"connDot\").className   = \"conn-dot \" + (ok ? \"ok\" : \"err\");\n  byId(\"connLabel\").textContent = ok ? \"APITable \u5df2\u8fde\u63a5\" : \"\u8fde\u63a5\u5931\u8d25\";\n}\n\nasync function api(path, body) {\n  const opts = {headers:{\"Content-Type\":\"application/json\"}};\n  if (body !== undefined) { opts.method = \"POST\"; opts.body = JSON.stringify(body); }\n  const r = await fetch(path, opts);\n  return r.json();\n}\n\n// \u2500\u2500 crowdin toggle \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\nfunction setCrowdin(yes) {\n  useCrowdin = yes;\n  byId(\"togYes\").classList.toggle(\"on\", yes);\n  byId(\"togNo\").classList.toggle(\"on\", !yes);\n\n  if (yes) {\n    // \u300c\u662f\u300d\uff1a\u663e\u793a\u4e0a\u4f20\u533a\uff0c\u94fe\u63a5\u53ea\u8bfb\u9884\u586b\uff0c\u9690\u85cf\u624b\u586b\u884c\n    byId(\"uploadSection\").style.display = \"block\";\n    byId(\"richTextSection\").style.display = \"block\";\n    byId(\"linkRow\").style.display = \"none\";\n    byId(\"f_link\").value = CROWDIN_DEFAULT;\n    updateFolderLabel();\n  } else {\n    // \u300c\u5426\u300d\uff1a\u9690\u85cf\u4e0a\u4f20\u533a\uff0c\u663e\u793a\u624b\u586b\u94fe\u63a5\u884c\uff0c\u6e05\u7a7a\u6587\u4ef6\n    byId(\"uploadSection\").style.display = \"none\";\n    byId(\"richTextSection\").style.display = \"none\";\n    byId(\"linkRow\").style.display = \"block\";\n    byId(\"f_link\").value = \"\";\n    byId(\"f_link\").readOnly = false;\n    fileBytes = null; fileName = \"\";\n    byId(\"dropName\").style.display = \"none\";\n    byId(\"dropOk\").style.display = \"none\";\n    byId(\"dropSize\").textContent = \"\";\n  }\n}\n\n// \u2500\u2500 file drop \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\nconst drop = byId(\"drop\");\ndrop.addEventListener(\"dragover\", e => { e.preventDefault(); drop.classList.add(\"over\"); });\ndrop.addEventListener(\"dragleave\", () => drop.classList.remove(\"over\"));\ndrop.addEventListener(\"drop\", e => {\n  e.preventDefault(); drop.classList.remove(\"over\");\n  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);\n});\nfunction onFile(inp) { if (inp.files.length) handleFile(inp.files[0]); }\nfunction handleFile(f) {\n  fileName = f.name;\n  const r = new FileReader(); r.onload = e => {\n    fileBytes = Array.from(new Uint8Array(e.target.result));\n    byId(\"dropName\").textContent = f.name; byId(\"dropName\").style.display = \"block\";\n    byId(\"dropOk\").style.display = \"inline-flex\";\n    byId(\"dropSize\").textContent = fmtB(f.size);\n  };\n  r.readAsArrayBuffer(f);\n\n  // \u81ea\u52a8\u628a\u6587\u4ef6\u540d\uff08\u53bb\u6389\u540e\u7f00\uff09\u586b\u5165\u7ffb\u8bd1\u9700\u6c42\u540d\u79f0\uff08\u5982\u679c\u5f53\u524d\u4e3a\u7a7a\uff09\n  const titleEl = byId(\"f_title\");\n  if (titleEl && !titleEl.value.trim()) {\n    const nameNoExt = f.name.replace(/\\.[^.]+$/, \"\");\n    const today2 = new Date();\n    const pad2 = n => String(n).padStart(2,\"0\");\n    const ds2 = `${today2.getFullYear()}${pad2(today2.getMonth()+1)}${pad2(today2.getDate())}`;\n    const req2 = v(\"f_req\");\n    titleEl.value = req2 ? `${ds2} ${nameNoExt} ${req2}` : `${ds2} ${nameNoExt}`;\n  }\n}\n\nfunction updateFolderLabel() {\n  const folder = v(\"crowdin_folder\") || \"English Team\";\n  const lbl = byId(\"folderLabel\");\n  if (lbl) lbl.textContent = folder;\n}\nfunction fmtB(b) {\n  if (b<1024) return b+\" B\";\n  if (b<1048576) return (b/1024).toFixed(1)+\" KB\";\n  return (b/1048576).toFixed(2)+\" MB\";\n}\n\n// \u2500\u2500 submit \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\nasync function doSubmit() {\n  // \u9a8c\u8bc1\n  if (!v(\"f_title\").trim())   { alert(\"\u8bf7\u586b\u5199\u7ffb\u8bd1\u9700\u6c42\u540d\u79f0\"); byId(\"f_title\").focus(); return; }\n  if (!v(\"f_req\"))            { alert(\"\u8bf7\u9009\u62e9\u9700\u6c42\u4eba\"); return; }\n  if (useCrowdin === null)    { alert(\"\u8bf7\u9009\u62e9\u300c\u662f\u5426\u7528\u8fd0\u8425\u7ec4\u4e13\u7528 Crowdin\u300d\"); return; }\n  if (!useCrowdin && !v(\"f_link\").trim()) { alert(\"\u9009\u300c\u5426\u300d\u65f6\u8bf7\u586b\u5199\u5b9e\u9645 Crowdin \u94fe\u63a5\"); return; }\n  const richHtml = byId(\"richEditor\") ? byId(\"richEditor\").innerHTML.trim() : \"\";\n  const hasRichContent = richHtml && richHtml !== \"\" && richHtml !== \"<br>\" && richHtml !== \"<div><br></div>\";\n  if (useCrowdin && !hasRichContent && (!fileBytes || !fileName)) { alert(\"\u8bf7\u9009\u62e9\u8981\u4e0a\u4f20\u7684\u6587\u4ef6\uff0c\u6216\u5728\u201c\u7eaf\u6587\u672c\u5185\u5bb9\u201d\u533a\u586b\u5199\u5185\u5bb9\"); return; }\n\n  // \u4fdd\u5b58\u914d\u7f6e\uff08\u542b Crowdin \u4e0a\u4f20\u8bbe\u7f6e\uff09\n  await api(\"/api/state\", {\n    current_user:       v(\"current_user\"),\n    crowdin_token:      v(\"crowdin_token\"),\n    apitable_api_key:   v(\"apitable_api_key\"),\n    crowdin_folder:     v(\"crowdin_folder\") || \"English Team\",\n    extra_path_keyword: v(\"extra_path_keyword\"),\n    extra_keyword:      v(\"extra_keyword\"),\n    save_crowdin:       byId(\"save_crowdin\").checked,\n    save_aitable:       byId(\"save_aitable\").checked,\n  });\n\n  // \u6784\u5efa APITable \u5b57\u6bb5\n  const fields = {};\n  fields[\"\\u7ffb\\u8bd1\\u9700\\u6c42\\uff08\\u5f53\\u65e5\\u65e5\\u671f+\\u7ffb\\u8bd1\\u9700\\u6c42\\u540d\\u5b57+\\u9700\\u6c42\\u4eba\\uff09\"] = v(\"f_title\").trim();\n  fields[\"\\u9700\\u6c42\\u4eba\"] = v(\"f_req\");\n  if (v(\"f_priority\")) fields[\"\\u4f18\\u5148\\u7ea7\"] = v(\"f_priority\");\n  if (v(\"f_ddl\"))      fields[\"\\u9700\\u6c42ddl\\uff08\\u5c3d\\u91cf\\u63d0\\u524d1-3\\u5c0f\\u65f6\\uff09\"] = new Date(v(\"f_ddl\")).getTime();\n  fields[\"\\u662f\\u5426\\u7528\\u8fd0\\u8425\\u7ec4\\u4e13\\u7528crowdin\"] = useCrowdin ? \"\\u662f\" : \"\\u5426\";\n  if (v(\"f_link\").trim()) fields[\"\\u586b\\u5199\\u5b9e\\u9645\\u4efb\\u52a1\\u94fe\\u63a5\\uff0c\\u9ed8\\u8ba4crowdin\\u94fe\\u63a5\"] = v(\"f_link\").trim();\n  if (v(\"f_note\").trim()) fields[\"\\u5907\\u6ce8(\\u4e0d\\u653e\\u5177\\u4f53\\u94fe\\u63a5)\"] = v(\"f_note\").trim();\n\n  byId(\"progressCard\").style.display = \"block\";\n  byId(\"progressCard\").scrollIntoView({behavior:\"smooth\"});\n  setSt(\"run\",\"\u23f3 \u63d0\u4ea4\u4e2d\u2026\");\n  byId(\"logbox\").textContent = \"\";\n  byId(\"resCard\").style.display = \"none\";\n  byId(\"submitBtn\").disabled = true;\n\n  // \u6587\u4ef6\u540d\uff1a\u5982\u679c\u6709\u9644\u4ef6\uff0c\u7528\u7ffb\u8bd1\u9700\u6c42\u540d\u79f0 + \u539f\u540e\u7f00\uff1b\u5982\u679c\u662f\u7eaf\u6587\u672c\uff0c\u7528\u7ffb\u8bd1\u9700\u6c42\u540d\u79f0 + .docx\n  let finalFilename = \"\";\n  let submitRichHtml = null;\n  if (hasRichContent && !fileBytes) {\n    // \u7eaf\u6587\u672c\u6a21\u5f0f\uff1a\u53d1\u9001 HTML \u5230\u540e\u7aef\u751f\u6210 docx\n    finalFilename = v(\"f_title\").trim() + \".docx\";\n    submitRichHtml = richHtml;\n  } else if (fileBytes && fileName) {\n    // \u6587\u4ef6\u6a21\u5f0f\uff1a\u7528\u7ffb\u8bd1\u9700\u6c42\u540d\u79f0 + \u539f\u6587\u4ef6\u540e\u7f00\n    const ext = fileName.includes(\".\") ? \".\" + fileName.split(\".\").pop() : \"\";\n    finalFilename = v(\"f_title\").trim() + ext;\n  }\n\n  try {\n    const res = await api(\"/api/submit\", {\n      apitable_api_key:   v(\"apitable_api_key\"),\n      crowdin_token:      v(\"crowdin_token\"),\n      use_crowdin:        useCrowdin,\n      upload_folder:      v(\"crowdin_folder\") || \"English Team\",\n      extra_path_keyword: v(\"extra_path_keyword\"),\n      extra_keyword:      v(\"extra_keyword\"),\n      fields:             fields,\n      filename:           finalFilename,\n      file_bytes:         fileBytes,\n      rich_html:          submitRichHtml,\n    });\n    if (res.error) throw new Error(res.error);\n    jobId = res.job_id;\n    if (poll) clearInterval(poll);\n    poll = setInterval(pollStatus, 1000);\n  } catch(e) {\n    setSt(\"err\",\"\u274c \" + e.message);\n    byId(\"submitBtn\").disabled = false;\n  }\n}\n\nasync function pollStatus() {\n  try {\n    const res = await api(\"/api/status?job_id=\"+jobId);\n    byId(\"logbox\").textContent = (res.log||[]).join(\"\\n\");\n    byId(\"logbox\").scrollTop = 99999;\n    if (res.status === \"success\") {\n      clearInterval(poll);\n      setSt(\"ok\",\"\u2705 \u63d0\u4ea4\u6210\u529f\uff01\");\n      byId(\"submitBtn\").disabled = false;\n      showResult(res.result);\n    } else if (res.status === \"error\") {\n      clearInterval(poll);\n      setSt(\"err\",\"\u274c \" + (res.error||\"\u5931\u8d25\"));\n      byId(\"submitBtn\").disabled = false;\n    }\n  } catch(_) {}\n}\n\nfunction showResult(r) {\n  if (!r) return;\n  const card = byId(\"resCard\");\n  card.style.display = \"block\";\n  byId(\"rRecId\").textContent = r.record_id || \"-\";\n  if (r.file_url) {\n    byId(\"rFileRow\").style.display = \"flex\";\n    const a = byId(\"rFileLink\");\n    a.href = r.file_url;\n    a.textContent = r.file_path || r.filename || r.file_url;\n  }\n}\n\nfunction setSt(type, msg) {\n  const el = byId(\"statusBar\");\n  el.className = \"status \" + ({run:\"s-run\",ok:\"s-ok\",err:\"s-err\"}[type]||\"s-idle\");\n  el.textContent = msg;\n}\n\nfunction resetForm() {\n  // \u53ea\u91cd\u7f6e\u4efb\u52a1\u8868\u5355\u5b57\u6bb5\uff0c\u4e0d\u91cd\u7f6e\u51ed\u8bc1\u8bbe\u7f6e\n  [\"f_title\",\"f_link\",\"f_note\",\"f_ddl\"].forEach(id => {\n    const el = byId(id); if (el) el.value=\"\";\n  });\n  byId(\"f_req\").value = \"\"; byId(\"f_priority\").value = \"\";\n  useCrowdin = null;\n  byId(\"togYes\").classList.remove(\"on\");\n  byId(\"togNo\").classList.remove(\"on\");\n  byId(\"f_link\").readOnly = false;\n  byId(\"uploadSection\").style.display = \"none\";\n  byId(\"richTextSection\").style.display = \"none\";\n  byId(\"linkRow\").style.display = \"none\";\n  fileBytes = null; fileName = \"\";\n  byId(\"dropName\").style.display = \"none\";\n  byId(\"dropOk\").style.display = \"none\";\n  byId(\"dropSize\").textContent = \"\";\n  byId(\"progressCard\").style.display = \"none\";\n  byId(\"submitBtn\").disabled = false;\n  const re = byId(\"richEditor\"); if (re) re.innerHTML = \"\";\n}"
    js_code = _JS.replace("__CROWDIN_URL__", crowdin_url)

    html_tmpl = (
        f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>翻译任务提交</title>
<style>
:root{{
  --bg:#0d1117;--card:#161b22;--border:#30363d;--text:#c9d1d9;
  --muted:#8b949e;--dim:#484f58;
  --blue:#1f6feb;--blue-h:#388bfd;
  --green:#238636;--green-h:#2ea043;
  --red:#da3633;--teal:#0ea5e9;
  --r:10px;
}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
     background:var(--bg);color:var(--text);min-height:100vh;padding-bottom:60px}}

.hdr{{background:var(--card);border-bottom:1px solid var(--border);
      padding:14px 28px;display:flex;align-items:center;gap:12px;
      position:sticky;top:0;z-index:100}}
.hdr-icon{{font-size:22px}}
.hdr-title{{font-size:15px;font-weight:700;color:#f0f6fc}}
.conn-dot{{width:8px;height:8px;border-radius:50%;background:var(--dim);
           margin-left:auto;flex-shrink:0;transition:.3s}}
.conn-dot.ok{{background:#3fb950}}
.conn-dot.err{{background:var(--red)}}
.conn-label{{font-size:11px;color:var(--muted)}}

.wrap{{max-width:720px;margin:0 auto;padding:28px 18px;display:grid;gap:20px}}

.card{{background:var(--card);border:1px solid var(--border);
       border-radius:var(--r);padding:22px}}
.card-title{{font-size:11px;font-weight:700;text-transform:uppercase;
             letter-spacing:.07em;color:var(--muted);margin-bottom:18px;
             display:flex;align-items:center;gap:8px}}
.card-title span{{color:#f0f6fc;font-size:13px;text-transform:none;
                   letter-spacing:0;font-weight:600}}
.card-title .spacer{{flex:1}}
.config-toggle{{border:1px solid var(--border);border-radius:999px;background:#0d1117;
  color:var(--text);font:inherit;font-size:12px;font-weight:700;padding:5px 10px;
  cursor:pointer;display:inline-flex;align-items:center;gap:6px;text-transform:none;letter-spacing:0}}
.config-dot{{width:8px;height:8px;border-radius:50%;background:var(--red);display:inline-block}}
.config-dot.ok{{background:var(--green)}}
.config-card.collapsed .config-body{{display:none}}
.config-card.collapsed{{padding-bottom:14px}}

.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}}
.task-meta{{display:grid;grid-template-columns:minmax(132px,150px) minmax(86px,96px) minmax(340px,1fr);gap:14px;align-items:end}}
@media(max-width:760px){{.task-meta{{grid-template-columns:1fr}}}}
@media(max-width:600px){{.g2,.g3{{grid-template-columns:1fr}}}}

.field{{display:flex;flex-direction:column;gap:6px}}
.field label{{font-size:12px;color:var(--muted);font-weight:600;
              display:flex;align-items:center;gap:4px}}
.field label .req{{color:var(--red);font-size:13px;line-height:1}}
.field input,.field select,.field textarea{{
  background:#0d1117;border:1px solid var(--border);border-radius:8px;
  padding:9px 12px;color:var(--text);font:inherit;font-size:13px;
  outline:none;transition:border-color .15s;width:100%}}
.field input:focus,.field select:focus,.field textarea:focus{{border-color:var(--blue)}}
.field textarea{{resize:vertical;min-height:72px}}
.field .hint{{font-size:11px;color:var(--muted);line-height:1.5;margin-top:2px}}
.rich-editor{{
  min-height:140px;background:#0d1117;border:1px solid var(--border);
  border-radius:8px;padding:10px 12px;color:var(--text);font:inherit;
  font-size:13px;line-height:1.6;outline:none;overflow:auto}}
.rich-editor:focus{{border-color:var(--blue)}}
.rich-editor:empty:before{{
  content:"在这里输入或粘贴要翻译的内容";color:var(--dim);pointer-events:none}}
.field input.locked{{background:#090d13;color:var(--muted);border-color:#21262d;cursor:not-allowed}}
.input-tabs{{display:flex;gap:10px;margin-bottom:12px}}
.tab-btn{{flex:1;padding:9px 12px;border-radius:8px;border:1px solid var(--border);
  background:#0d1117;color:var(--muted);font:inherit;font-size:13px;font-weight:700;
  cursor:pointer;transition:.15s}}
.tab-btn.on{{border-color:var(--blue);color:#c9d1d9;background:#111827}}
.mode-pane{{display:none}}
.mode-pane.on{{display:block}}
.file-meta{{display:flex;gap:10px;align-items:center;justify-content:center;flex-wrap:wrap;margin-top:10px;
  position:relative;z-index:3}}
.mini-btn{{border:1px solid var(--border);border-radius:999px;background:#21262d;color:var(--text);
  font:inherit;font-size:12px;font-weight:700;padding:5px 10px;cursor:pointer;position:relative;z-index:4}}
.mini-btn:hover{{background:#30363d}}
.ddl-row{{display:grid;grid-template-columns:minmax(0,1fr) 108px 44px;gap:10px;align-items:end}}
.ddl-date-wrap{{position:relative;display:flex;align-items:stretch}}
.ddl-date-native{{position:absolute;width:1px;height:1px;opacity:0;pointer-events:none}}
.ddl-date-control{{width:100%;height:39px;border:1px solid var(--border);border-radius:8px;
  background:#0d1117;color:var(--text);cursor:pointer;font:inherit;font-size:13px;
  display:flex;align-items:center;justify-content:space-between;gap:8px;padding:9px 12px;
  transition:border-color .15s,background .15s}}
.ddl-date-control:hover,.ddl-date-control:focus{{border-color:var(--blue);background:#111827;outline:none}}
.ddl-date-control span:first-child{{white-space:nowrap}}
.ddl-caret{{font-size:14px;line-height:1;color:var(--muted)}}
.ddl-hour{{width:108px;min-width:108px;text-align:center;text-align-last:center;appearance:auto}}
.ddl-fixed{{
  height:39px;background:#0d1117;border:1px solid var(--border);border-radius:8px;
  padding:9px 12px;color:var(--text);font:inherit;font-size:13px;display:flex;
  align-items:center;justify-content:center;letter-spacing:0}}

/* crowdin toggle */
.toggle-row{{display:flex;gap:10px;margin-bottom:4px}}
.tog-btn{{flex:1;padding:10px;border-radius:8px;border:1.5px solid var(--border);
          background:#0d1117;color:var(--muted);font:inherit;font-size:13px;
          font-weight:600;cursor:pointer;transition:.15s;text-align:center}}
.tog-btn.yes.on{{border-color:#3fb950;background:#122d20;color:#3fb950}}
.tog-btn.no.on{{border-color:#f0883e;background:#2d1a10;color:#f0883e}}
.tog-btn:not(.on):hover{{border-color:var(--blue);color:var(--text)}}

/* file drop */
.drop{{border:2px dashed var(--border);border-radius:var(--r);padding:32px 20px;
       text-align:center;cursor:pointer;transition:.18s;position:relative;background:#0a0d13}}
.drop:hover,.drop.over{{border-color:var(--blue);background:#0c1525}}
.drop input[type=file]{{position:absolute;inset:0;opacity:0;cursor:pointer;
                         width:100%;height:100%;z-index:1;pointer-events:none}}
.drop-icon{{font-size:36px;margin-bottom:8px;position:relative;z-index:2}}
.drop-txt{{font-size:13px;color:var(--muted);position:relative;z-index:2}}
.drop-name{{font-size:13px;font-weight:600;color:#f0f6fc;margin-top:8px;word-break:break-all;
  position:relative;z-index:2}}
.drop-ok{{display:inline-flex;align-items:center;gap:5px;background:#122d20;
           color:#3fb950;padding:4px 11px;border-radius:999px;font-size:12px;
           font-weight:700;margin-top:6px}}
.drop-size{{font-size:11px;color:var(--muted);margin-top:3px;position:relative;z-index:2}}
.file-hidden{{display:none}}

/* submit */
.submit-row{{display:flex;gap:12px;align-items:center;flex-wrap:wrap}}
.btn{{border:none;border-radius:8px;padding:10px 22px;font:inherit;
      font-size:14px;font-weight:700;cursor:pointer;transition:.14s}}
.btn:disabled{{opacity:.4;cursor:not-allowed}}
.btn-green{{background:var(--green);color:#fff}}
.btn-green:hover:not(:disabled){{background:var(--green-h)}}
.btn-ghost{{background:#21262d;border:1px solid var(--border);color:var(--text)}}
.btn-ghost:hover:not(:disabled){{background:#30363d}}

/* status & log */
.status{{padding:11px 15px;border-radius:8px;font-size:13px;font-weight:600;margin-bottom:10px}}
.s-idle{{background:#21262d;color:var(--muted)}}
.s-run{{background:#1c2a3a;color:#79c0ff}}
.s-ok{{background:#122d20;color:#3fb950}}
.s-err{{background:#32100b;color:#f85149}}
.logbox{{background:#080d14;border:1px solid #21262d;border-radius:var(--r);
         padding:14px 16px;min-height:160px;max-height:300px;overflow-y:auto;
         font-family:"SF Mono",Menlo,Consolas,monospace;font-size:12px;
         color:#d7e3f4;line-height:1.75;white-space:pre-wrap}}

/* result */
.res-card{{background:#0d1117;border:1px solid #238636;border-radius:var(--r);
           padding:16px;margin-top:14px;display:none}}
.res-card a{{color:#58a6ff;text-decoration:none}}
.res-card a:hover{{text-decoration:underline}}
.res-row{{display:flex;gap:8px;margin-top:8px;font-size:13px}}
.res-k{{color:var(--muted);min-width:100px;flex-shrink:0}}
.res-v{{color:var(--text);word-break:break-all}}

.chk{{display:flex;align-items:center;gap:7px;font-size:12px;color:var(--muted);
      cursor:pointer;margin-top:6px}}
.chk input{{width:14px;height:14px;accent-color:var(--blue);cursor:pointer}}

.spin{{display:inline-block;animation:spin .7s linear infinite}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
</style>
</head>
<body>

<div class="hdr">
  <div class="hdr-icon">📋</div>
  <div class="hdr-title">翻译任务提交</div>
  <div class="conn-dot" id="connDot"></div>
  <div class="conn-label" id="connLabel">连接中…</div>
</div>

<div class="wrap">

  <!-- 设置 -->
  <div class="card config-card" id="configCard">
    <div class="card-title">⚙ <span>凭证设置</span><span class="spacer"></span><button type="button" class="config-toggle" onclick="openConfig()"><span class="config-dot" id="configDot"></span><span id="configBadge">待配置</span></button></div>
    <div class="config-body" id="configBody">
    <div class="g2" style="margin-bottom:12px">
      <div class="field">
        <label>使用人 <span class="req">*</span></label>
        <select id="current_user">
          <option value="">-- 选择名字 --</option>
          {req_options}
        </select>
      </div>
      <div class="field">
        <label>Crowdin API Token <span class="req">*</span></label>
        <input id="crowdin_token" type="password" placeholder="上传附件时需要">
      </div>
    </div>
    <div class="field" style="margin-bottom:12px">
      <label>APITable API Key <span class="req">*</span></label>
      <input id="apitable_api_key" type="password" placeholder="Bearer Key">
    </div>
    <!-- Crowdin 上传设置 -->
    <div class="g3" style="margin-bottom:12px">
      <div class="field">
        <label>Crowdin 上传文件夹</label>
        <input id="crowdin_folder" type="text" placeholder="English Team" value="English Team">
      </div>
      <div class="field">
        <label>附加路径关键词</label>
        <input id="extra_path_keyword" type="text" placeholder="可选，如 activity/mail">
      </div>
      <div class="field">
        <label>附加文件关键词</label>
        <input id="extra_keyword" type="text" placeholder="可选，缩小匹配范围">
      </div>
    </div>
    <div class="field" style="margin-bottom:12px">
      <label>Crowdin 翻译语言</label>
      <div class="ddl-fixed" style="justify-content:flex-start">en-US</div>
      <div class="hint">默认目标语言，脚本内固定为 en-US</div>
    </div>
    <div style="display:flex;gap:20px;flex-wrap:wrap;align-items:center">
      <label class="chk"><input id="save_crowdin" type="checkbox" checked>保存 Crowdin Token</label>
      <label class="chk"><input id="save_aitable" type="checkbox" checked>保存 APITable Key</label>
      <button type="button" class="btn btn-ghost" onclick="saveSettings()">💾 保存设置</button>
    </div>
    </div>
  </div>

  <!-- 新建任务表单 -->
  <div class="card" id="taskCard">
    <div class="card-title">📝 <span>新建翻译任务</span></div>

    <!-- 翻译需求名称 -->
    <div class="field" style="margin-bottom:14px">
      <label>翻译需求名称 <span class="req">*</span></label>
      <input id="f_title" type="text"
        placeholder="选择附件时自动填充文件名；文本录入时，需要创建名字">
    </div>

    <div class="task-meta" style="margin-bottom:14px">
      <div class="field">
        <label>需求人 <span class="req">*</span></label>
        <select id="f_req">
          <option value="">-- 选择 --</option>
          {req_options}
        </select>
      </div>
      <div class="field">
        <label>优先级</label>
        <select id="f_priority">
          <option value="">-- 选择 --</option>
          <option>S</option><option>A</option><option selected>B</option><option>C</option>
        </select>
      </div>
      <div class="field">
        <label>需求 DDL</label>
        <div class="ddl-row">
          <div class="ddl-date-wrap">
            <input id="f_ddl_date" class="ddl-date-native" type="date" onchange="updateDdlDateLabel()">
            <button type="button" class="ddl-date-control" onclick="openDatePicker()" aria-label="打开日期选择">
              <span id="f_ddl_date_label">-- 日期 --</span><span class="ddl-caret">▾</span>
            </button>
          </div>
          <select id="f_ddl_hour" class="ddl-hour">
            <option value="">-- 小时 --</option>
            <option value="00">00</option><option value="01">01</option><option value="02">02</option><option value="03">03</option>
            <option value="04">04</option><option value="05">05</option><option value="06">06</option><option value="07">07</option>
            <option value="08">08</option><option value="09">09</option><option value="10">10</option><option value="11">11</option>
            <option value="12">12</option><option value="13">13</option><option value="14">14</option><option value="15">15</option>
            <option value="16">16</option><option value="17">17</option><option value="18">18</option><option value="19">19</option>
            <option value="20">20</option><option value="21">21</option><option value="22">22</option><option value="23">23</option>
          </select>
          <div class="ddl-fixed">00</div>
        </div>
      </div>
    </div>

    <!-- 备注 -->
    <div class="field" style="margin-bottom:14px">
      <label>备注（不放具体链接）</label>
      <textarea id="f_note" placeholder="可选"></textarea>
    </div>

    <!-- 是否用运营组 Crowdin -->
    <div class="field" style="margin-bottom:14px">
      <label>是否用运营组专用 Crowdin <span class="req">*</span></label>
      <div class="toggle-row">
        <button type="button" class="tog-btn yes" id="togYes" onclick="setCrowdin(true)">✅ 是（运营组 Crowdin）</button>
        <button type="button" class="tog-btn no"  id="togNo"  onclick="setCrowdin(false)">❌ 否（填写实际链接）</button>
      </div>
    </div>

    <!-- Crowdin 链接（「是」时只读预填，「否」时手填） -->
    <div class="field" id="linkRow" style="margin-bottom:14px;display:none">
      <label id="linkLabel">实际任务链接</label>
      <input id="f_link" type="text" placeholder="请填写实际 Crowdin 链接">
    </div>

    <!-- 附件 / 文本录入（二选一） -->
    <div id="uploadSection" style="margin-bottom:14px;display:none">
      <div class="input-tabs">
        <button type="button" class="tab-btn on" id="tabFile" onclick="setInputMode('file')">附件</button>
        <button type="button" class="tab-btn" id="tabText" onclick="setInputMode('text')">文本录入</button>
      </div>
      <div class="mode-pane on" id="filePane">
        <div class="field" style="margin-bottom:10px">
          <label>📎 上传附件到 Crowdin <span id="folderLabel" style="color:var(--teal);font-weight:700">English Team</span></label>
          <div class="drop" id="drop">
            <input type="file" id="fileInput" onclick="this.value=''" onchange="onFile(this)">
            <div class="drop-icon">📁</div>
            <div class="drop-txt">点击选择 / 拖拽文件到这里</div>
            <div id="dropName" style="display:none" class="drop-name"></div>
            <div class="file-meta">
              <div id="dropOk" style="display:none" class="drop-ok">✓ 已选择</div>
              <button type="button" id="removeFileBtn" class="mini-btn" style="display:none" onclick="event.stopPropagation(); removeAttachment()">去除附件</button>
            </div>
            <div id="dropSize" class="drop-size"></div>
          </div>
        </div>
      </div>
      <div class="mode-pane" id="textPane">
        <div class="field">
          <label>✏️ 文本录入（提交时自动生成 .docx 上传）</label>
          <div class="hint">文本录入时需要手动填写翻译需求名称。文件名 = 翻译需求名称.docx</div>
          <div id="richEditor" contenteditable="true" class="rich-editor"></div>
        </div>
      </div>
    </div>

    <!-- 提交 -->
    <div class="submit-row">
      <button type="button" class="btn btn-green" id="submitBtn" onclick="doSubmit()">🚀 提交任务</button>
      <button type="button" class="btn btn-ghost" onclick="resetForm()">↺ 重置</button>
    </div>
  </div>

  <!-- 进度 & 结果 -->
  <div class="card" id="progressCard" style="display:none">
    <div class="card-title">📊 <span>提交进度</span></div>
    <div id="statusBar" class="status s-idle">等待中</div>
    <div id="logbox" class="logbox"></div>
    <div class="res-card" id="resCard">
      <div style="font-size:13px;font-weight:700;color:#3fb950;margin-bottom:4px">✅ 提交成功</div>
      <div class="res-row"><div class="res-k">记录 ID</div><div class="res-v" id="rRecId">-</div></div>
      <div class="res-row" id="rFileRow" style="display:none">
        <div class="res-k">Crowdin 文件</div>
        <div class="res-v"><a id="rFileLink" href="#" target="_blank">-</a></div>
      </div>
    </div>
  </div>

</div><!-- /wrap -->

<script>"""
        + js_code +
        """</script>
</body>
</html>"""
    )
    html_tmpl = html_tmpl.replace(
        '// ── 关闭浏览器时自动关闭服务 ─────────────────────────────────────────────────\nwindow.addEventListener("beforeunload", () => {\n  navigator.sendBeacon("/api/shutdown");\n});',
        'let heartbeatTimer = null;\n\nfunction sendHeartbeat() {\n  fetch("/api/heartbeat", {\n    method: "POST",\n    headers: {"Content-Type":"application/json"},\n    body: "{}",\n    keepalive: true,\n  }).catch(() => {});\n}\n\nfunction startHeartbeat() {\n  sendHeartbeat();\n  if (!heartbeatTimer) heartbeatTimer = setInterval(sendHeartbeat, 3000);\n}\n\nwindow.addEventListener("pageshow", startHeartbeat);\nwindow.addEventListener("focus", sendHeartbeat);\ndocument.addEventListener("visibilitychange", () => {\n  if (!document.hidden) sendHeartbeat();\n});',
        1,
    )
    html_tmpl = html_tmpl.replace(
        'window.addEventListener("load", async () => {\n  // 自动填今天日期到标题',
        'window.addEventListener("load", async () => {\n  startHeartbeat();\n  // 自动填今天日期到标题',
        1,
    )
    html_tmpl = html_tmpl.replace(
        'byId("f_title").placeholder = `格式：${ds} 翻译需求名字 需求人`;',
        'byId("f_title").placeholder = "选择附件时自动填充文件名；文本录入时，需要创建名字";',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '<option>S</option><option>A</option><option>B</option><option>C</option>',
        '<option>S</option><option>A</option><option selected>B</option><option>C</option>',
    )
    html_tmpl = html_tmpl.replace(
        '    </div>\n    <div style="display:flex;gap:20px;flex-wrap:wrap">',
        '    </div>\n    <div class="field" style="margin-bottom:12px">\n      <label>Crowdin 翻译语言</label>\n      <div class="ddl-fixed" style="justify-content:flex-start">en-US</div>\n      <div class="hint">默认目标语言，脚本内固定为 en-US</div>\n    </div>\n    <div style="display:flex;gap:20px;flex-wrap:wrap">',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '    <div class="field" style="margin-bottom:12px">\n      <label>Crowdin 翻译语言</label>\n      <input id="crowdin_language" type="text" value="en-US" placeholder="en-US">\n      <div class="hint">默认目标语言，直接使用 en-US</div>\n    </div>\n',
        '    <div class="field" style="margin-bottom:12px">\n      <label>Crowdin 翻译语言</label>\n      <div class="ddl-fixed" style="justify-content:flex-start">en-US</div>\n      <div class="hint">默认目标语言，脚本内固定为 en-US</div>\n    </div>\n',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '<input id="f_ddl" type="datetime-local">',
        '<div class="ddl-row">\n          <input id="f_ddl_date" type="date">\n          <select id="f_ddl_hour">\n            <option value="">-- 小时 --</option>\n            <option value="00">00</option><option value="01">01</option><option value="02">02</option><option value="03">03</option>\n            <option value="04">04</option><option value="05">05</option><option value="06">06</option><option value="07">07</option>\n            <option value="08">08</option><option value="09">09</option><option value="10">10</option><option value="11">11</option>\n            <option value="12">12</option><option value="13">13</option><option value="14">14</option><option value="15">15</option>\n            <option value="16">16</option><option value="17">17</option><option value="18">18</option><option value="19">19</option>\n            <option value="20">20</option><option value="21">21</option><option value="22">22</option><option value="23">23</option>\n          </select>\n          <div class="ddl-fixed">00</div>\n        </div>',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '[\"current_user\",\"crowdin_token\",\"apitable_api_key\",\n   \"crowdin_folder\",\"extra_path_keyword\",\"extra_keyword\"]',
        '[\"current_user\",\"crowdin_token\",\"apitable_api_key\",\n   \"crowdin_folder\",\"extra_path_keyword\",\"extra_keyword\"]',
    )
    html_tmpl = html_tmpl.replace(
        '  // 同步 req 与 current_user\n  if (d.current_user) byId(\"f_req\").value = d.current_user;\n',
        '  syncRequesterFromUser();\n',
    )
    html_tmpl = html_tmpl.replace(
        '    applyConfig(cfg);\n    // 静默检查连接',
        '    applyConfig(cfg);\n    updateConfigState(isConfigReady());\n    const userEl = byId(\"current_user\");\n    if (userEl) userEl.addEventListener(\"change\", syncRequesterFromUser);\n    // 静默检查连接',
        1,
    )
    html_tmpl = html_tmpl.replace(
        'function setConn(ok) {\n  byId(\"connDot\").className   = \"conn-dot \" + (ok ? \"ok\" : \"err\");\n  byId(\"connLabel\").textContent = ok ? \"APITable 已连接\" : \"连接失败\";\n}\n\nasync function api(path, body) {',
        'function setConn(ok) {\n  byId(\"connDot\").className   = \"conn-dot \" + (ok ? \"ok\" : \"err\");\n  byId(\"connLabel\").textContent = ok ? \"APITable 已连接\" : \"连接失败\";\n}\n\nfunction syncRequesterFromUser() {\n  const user = v(\"current_user\");\n  const reqEl = byId(\"f_req\");\n  if (reqEl) reqEl.value = user || \"\";\n}\n\nfunction saveSettings() {\n  return api(\"/api/state\", {\n    current_user:       v(\"current_user\"),\n    crowdin_token:      v(\"crowdin_token\"),\n    apitable_api_key:   v(\"apitable_api_key\"),\n    crowdin_folder:     v(\"crowdin_folder\") || \"English Team\",\n    extra_path_keyword: v(\"extra_path_keyword\"),\n    extra_keyword:      v(\"extra_keyword\"),\n    save_crowdin:       true,\n    save_aitable:       true,\n  }).then(() => {\n    byId(\"save_crowdin\").checked = true;\n    byId(\"save_aitable\").checked = true;\n    alert(\"设置已保存\");\n  });\n}\n\nfunction openDatePicker() {\n  const el = byId(\"f_ddl_date\");\n  if (el && el.showPicker) el.showPicker();\n  else if (el) el.focus();\n}\n\nfunction updateDdlDateLabel() {\n  const el = byId(\"f_ddl_date\");\n  const label = byId(\"f_ddl_date_label\");\n  if (label) label.textContent = el && el.value ? el.value : \"-- 日期 --\";\n}\n\nasync function api(path, body) {',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '  // 自动把文件名（去掉后缀）填入翻译需求名称（如果当前为空）\n  const titleEl = byId(\"f_title\");\n  if (titleEl && !titleEl.value.trim()) {\n    const nameNoExt = f.name.replace(/\\.[^.]+$/, \"\");\n    const today2 = new Date();\n    const pad2 = n => String(n).padStart(2,\"0\");\n    const ds2 = `${today2.getFullYear()}${pad2(today2.getMonth()+1)}${pad2(today2.getDate())}`;\n    const req2 = v(\"f_req\");\n    titleEl.value = req2 ? `${ds2} ${nameNoExt} ${req2}` : `${ds2} ${nameNoExt}`;\n  }\n',
        '  // 自动把文件名（去掉后缀）直接填入翻译需求名称\n  const titleEl = byId(\"f_title\");\n  if (titleEl) titleEl.value = f.name.replace(/\\.[^.]+$/, \"\");\n',
        1,
    )
    html_tmpl = html_tmpl.replace(
        'function fmtB(b) {\n  if (b<1024) return b+\" B\";\n  if (b<1048576) return (b/1024).toFixed(1)+\" KB\";\n  return (b/1048576).toFixed(2)+\" MB\";\n}\n\n// ── submit ────────────────────────────────────────────────────────────────────',
        'function fmtB(b) {\n  if (b<1024) return b+\" B\";\n  if (b<1048576) return (b/1024).toFixed(1)+\" KB\";\n  return (b/1048576).toFixed(2)+\" MB\";\n}\n\nfunction buildDdlValue() {\n  const date = v(\"f_ddl_date\");\n  const hour = v(\"f_ddl_hour\");\n  if (!date || hour === \"\") return \"\";\n  return `${date}T${hour}:00`;\n}\n\n// ── submit ────────────────────────────────────────────────────────────────────',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '  if (!v(\"f_req\"))            { alert(\"请选择需求人\"); return; }\n',
        '  syncRequesterFromUser();\n  if (!v(\"current_user\"))     { alert(\"请选择使用人\"); return; }\n',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '    crowdin_folder:     v(\"crowdin_folder\") || \"English Team\",\n    extra_path_keyword: v(\"extra_path_keyword\"),',
        '    crowdin_folder:     v(\"crowdin_folder\") || \"English Team\",\n    crowdin_language:   v(\"crowdin_language\") || \"en-US\",\n    extra_path_keyword: v(\"extra_path_keyword\"),',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '  fields[\"\\u9700\\u6c42\\u4eba\"] = v(\"f_req\");\n  if (v(\"f_priority\")) fields[\"\\u4f18\\u5148\\u7ea7\"] = v(\"f_priority\");\n  if (v(\"f_ddl\"))      fields[\"\\u9700\\u6c42ddl\\uff08\\u5c3d\\u91cf\\u63d0\\u524d1-3\\u5c0f\\u65f6\\uff09\"] = new Date(v(\"f_ddl\")).getTime();\n',
        '  const requester = v(\"current_user\") || v(\"f_req\");\n  fields[\"\\u9700\\u6c42\\u4eba\"] = requester;\n  if (v(\"f_priority\")) fields[\"\\u4f18\\u5148\\u7ea7\"] = v(\"f_priority\");\n  const ddlValue = buildDdlValue();\n  if (ddlValue) fields[\"\\u9700\\u6c42ddl\\uff08\\u5c3d\\u91cf\\u63d0\\u524d1-3\\u5c0f\\u65f6\\uff09\"] = new Date(ddlValue).getTime();\n',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '      upload_folder:      v(\"crowdin_folder\") || \"English Team\",\n      extra_path_keyword: v(\"extra_path_keyword\"),',
        '      upload_folder:      v(\"crowdin_folder\") || \"English Team\",\n      crowdin_language:   \"en-US\",\n      extra_path_keyword: v(\"extra_path_keyword\"),',
        1,
    )
    html_tmpl = html_tmpl.replace(
        '  [\"f_title\",\"f_link\",\"f_note\",\"f_ddl\"].forEach(id => {\n    const el = byId(id); if (el) el.value=\"\";\n  });\n  byId(\"f_req\").value = \"\"; byId(\"f_priority\").value = \"\";\n',
        '  [\"f_title\",\"f_link\",\"f_note\",\"f_ddl_date\"].forEach(id => {\n    const el = byId(id); if (el) el.value=\"\";\n  });\n  updateDdlDateLabel();\n  byId(\"f_ddl_hour\").value = \"\";\n  syncRequesterFromUser();\n  byId(\"f_priority\").value = \"B\";\n',
        1,
    )
    extra_js = r'''
let inputMode = "file";
let titleLockedByAttachment = false;

async function api(path, body) {
  const opts = {headers:{"Content-Type":"application/json"}};
  if (body !== undefined) { opts.method = "POST"; opts.body = JSON.stringify(body); }
  const r = await fetch(path, opts);
  const data = await r.json().catch(() => ({}));
  if (!r.ok || data.error) throw new Error(data.error || `请求失败：HTTP ${r.status}`);
  return data;
}

function isConfigReady() {
  return !!(v("current_user") && v("crowdin_token").trim() && v("apitable_api_key").trim());
}

function openConfig() {
  const card = byId("configCard");
  if (card) card.classList.remove("collapsed");
}

function updateConfigState(collapseWhenReady=false) {
  const ready = isConfigReady();
  const dot = byId("configDot");
  const badge = byId("configBadge");
  const card = byId("configCard");
  if (dot) dot.classList.toggle("ok", ready);
  if (badge) badge.textContent = ready ? "已配置" : "待配置";
  if (card && ready && collapseWhenReady) card.classList.add("collapsed");
  if (card && !ready) card.classList.remove("collapsed");
  return ready;
}

function requireConfig() {
  if (isConfigReady()) return true;
  openConfig();
  updateConfigState(false);
  alert("请先填写使用人、Crowdin API Token 和 APITable API Key，并点击保存设置。");
  return false;
}

function saveSettings() {
  if (!v("current_user")) { alert("请先选择使用人"); byId("current_user").focus(); return Promise.resolve(false); }
  if (!v("crowdin_token").trim()) { alert("请填写 Crowdin API Token"); byId("crowdin_token").focus(); return Promise.resolve(false); }
  if (!v("apitable_api_key").trim()) { alert("请填写 APITable API Key"); byId("apitable_api_key").focus(); return Promise.resolve(false); }
  return api("/api/state", {
    current_user:       v("current_user"),
    crowdin_token:      v("crowdin_token"),
    apitable_api_key:   v("apitable_api_key"),
    crowdin_folder:     v("crowdin_folder") || "English Team",
    extra_path_keyword: v("extra_path_keyword"),
    extra_keyword:      v("extra_keyword"),
    save_crowdin:       true,
    save_aitable:       true,
  }).then(() => {
    byId("save_crowdin").checked = true;
    byId("save_aitable").checked = true;
    updateConfigState(true);
    alert("✅ 设置已保存，正在验证 APITable 连接…");
    checkConn(v("apitable_api_key"));
    return true;
  }).catch(e => {
    alert("❌ 保存设置失败：" + (e.message || e));
    return false;
  });
}

function setTitleLocked(locked) {
  titleLockedByAttachment = locked;
  const titleEl = byId("f_title");
  if (!titleEl) return;
  titleEl.readOnly = locked;
  titleEl.classList.toggle("locked", locked);
}

function setTitleFromFile(name) {
  const titleEl = byId("f_title");
  if (!titleEl) return;
  titleEl.value = name.replace(/\.[^.]+$/, "");
  setTitleLocked(true);
}

function clearRichText() {
  const re = byId("richEditor");
  if (re) re.innerHTML = "";
}

function hasRichText() {
  const re = byId("richEditor");
  const html = re ? re.innerHTML.trim() : "";
  return !!(html && html !== "<br>" && html !== "<div><br></div>");
}

function removeAttachment(clearTitle=true) {
  fileBytes = null;
  fileName = "";
  const input = byId("fileInput");
  if (input) input.value = "";
  byId("dropName").style.display = "none";
  byId("dropOk").style.display = "none";
  byId("removeFileBtn").style.display = "none";
  byId("dropSize").textContent = "";
  setTitleLocked(false);
  if (clearTitle) byId("f_title").value = "";
}

function openFileChooser() {
  const input = byId("fileInput");
  if (input) input.click();
}

function setInputMode(mode) {
  if (mode === inputMode) return true;
  if (mode === "text" && fileBytes) {
    alert("已选择的附件不会被上传，将切换为文本录入并清空附件；翻译需求名称会解锁并重置。");
    removeAttachment(true);
  }
  if (mode === "file" && hasRichText()) {
    alert("已填写的文本不会被上传，将切换为附件上传。选择附件后会自动填入文件名并锁定需求名称。");
    clearRichText();
  }
  inputMode = mode;
  byId("tabFile").classList.toggle("on", mode === "file");
  byId("tabText").classList.toggle("on", mode === "text");
  byId("filePane").classList.toggle("on", mode === "file");
  byId("textPane").classList.toggle("on", mode === "text");
  if (mode === "text") setTitleLocked(false);
  if (mode === "file" && fileName) setTitleFromFile(fileName);
  return true;
}

function setCrowdin(yes) {
  if (!requireConfig()) return;
  useCrowdin = yes;
  byId("togYes").classList.toggle("on", yes);
  byId("togNo").classList.toggle("on", !yes);
  if (yes) {
    byId("uploadSection").style.display = "block";
    byId("linkRow").style.display = "none";
    byId("f_link").value = CROWDIN_DEFAULT;
    setInputMode(inputMode || "file");
    updateFolderLabel();
  } else {
    byId("uploadSection").style.display = "none";
    byId("linkRow").style.display = "block";
    byId("f_link").value = "";
    byId("f_link").readOnly = false;
    removeAttachment(true);
    clearRichText();
  }
}

function handleFile(f) {
  if (hasRichText()) {
    alert("文本录入内容不会被上传，将切换为附件上传并清空文本。");
    clearRichText();
  }
  inputMode = "file";
  byId("tabFile").classList.add("on");
  byId("tabText").classList.remove("on");
  byId("filePane").classList.add("on");
  byId("textPane").classList.remove("on");
  fileName = f.name;
  const r = new FileReader();
  r.onload = e => {
    fileBytes = Array.from(new Uint8Array(e.target.result));
    byId("dropName").textContent = f.name;
    byId("dropName").style.display = "block";
    byId("dropOk").style.display = "inline-flex";
    byId("removeFileBtn").style.display = "inline-flex";
    byId("dropSize").textContent = fmtB(f.size);
  };
  r.readAsArrayBuffer(f);
  setTitleFromFile(f.name);
}

async function doSubmit() {
  if (!requireConfig()) return;
  if (!v("f_title").trim()) { alert("请填写翻译需求名称"); byId("f_title").focus(); return; }
  syncRequesterFromUser();
  if (!v("current_user")) { alert("请选择使用人"); return; }
  if (useCrowdin === null) { alert("请选择「是否用运营组专用 Crowdin」"); return; }
  if (!useCrowdin && !v("f_link").trim()) { alert("选「否」时请填写实际 Crowdin 链接"); return; }

  const richHtml = hasRichText() ? byId("richEditor").innerHTML.trim() : "";
  if (useCrowdin && inputMode === "file" && (!fileBytes || !fileName)) { alert("请选择要上传的附件"); return; }
  if (useCrowdin && inputMode === "text" && !richHtml) { alert("请在文本录入页签填写内容"); return; }

  // 先验证 APITable 连接
  try {
    const pingRes = await api("/api/ping", {apitable_api_key: v("apitable_api_key")});
    if (!pingRes.ok) {
      setConn(false);
      alert("❌ APITable 连接失败，请检查 API Key 是否正确。");
      return;
    }
    setConn(true);
  } catch(e) {
    setConn(false);
    alert("❌ APITable 连接验证出错：" + (e.message || e));
    return;
  }

  // 保存配置
  try {
    await api("/api/state", {
      current_user:       v("current_user"),
      crowdin_token:      v("crowdin_token"),
      apitable_api_key:   v("apitable_api_key"),
      crowdin_folder:     v("crowdin_folder") || "English Team",
      extra_path_keyword: v("extra_path_keyword"),
      extra_keyword:      v("extra_keyword"),
      save_crowdin:       true,
      save_aitable:       true,
    });
  } catch(e) {
    alert("❌ 保存配置失败：" + (e.message || e));
    return;
  }

  const fields = {};
  fields["\u7ffb\u8bd1\u9700\u6c42\uff08\u5f53\u65e5\u65e5\u671f+\u7ffb\u8bd1\u9700\u6c42\u540d\u5b57+\u9700\u6c42\u4eba\uff09"] = v("f_title").trim();
  fields["\u9700\u6c42\u4eba"] = v("current_user") || v("f_req");
  if (v("f_priority")) fields["\u4f18\u5148\u7ea7"] = v("f_priority");
  const ddlValue = buildDdlValue();
  if (ddlValue) fields["\u9700\u6c42ddl\uff08\u5c3d\u91cf\u63d0\u524d1-3\u5c0f\u65f6\uff09"] = new Date(ddlValue).getTime();
  fields["\u662f\u5426\u7528\u8fd0\u8425\u7ec4\u4e13\u7528crowdin"] = useCrowdin ? "\u662f" : "\u5426";
  if (v("f_link").trim()) fields["\u586b\u5199\u5b9e\u9645\u4efb\u52a1\u94fe\u63a5\uff0c\u9ed8\u8ba4crowdin\u94fe\u63a5"] = v("f_link").trim();
  if (v("f_note").trim()) fields["\u5907\u6ce8(\u4e0d\u653e\u5177\u4f53\u94fe\u63a5)"] = v("f_note").trim();

  byId("progressCard").style.display = "block";
  byId("progressCard").scrollIntoView({behavior:"smooth"});
  setSt("run","⏳ 提交中…");
  byId("logbox").textContent = "";
  byId("resCard").style.display = "none";
  byId("submitBtn").disabled = true;

  let finalFilename = "";
  let submitRichHtml = null;
  let submitFileBytes = null;
  if (useCrowdin && inputMode === "text") {
    finalFilename = v("f_title").trim() + ".docx";
    submitRichHtml = richHtml;
    submitFileBytes = null;
  } else if (useCrowdin && inputMode === "file") {
    const ext = fileName.includes(".") ? "." + fileName.split(".").pop() : "";
    finalFilename = v("f_title").trim() + ext;
    submitFileBytes = fileBytes;
  }

  try {
    const res = await api("/api/submit", {
      apitable_api_key:   v("apitable_api_key"),
      crowdin_token:      v("crowdin_token"),
      use_crowdin:        useCrowdin,
      upload_folder:      v("crowdin_folder") || "English Team",
      crowdin_language:   "en-US",
      extra_path_keyword: v("extra_path_keyword"),
      extra_keyword:      v("extra_keyword"),
      fields:             fields,
      filename:           finalFilename,
      file_bytes:         submitFileBytes,
      rich_html:          submitRichHtml,
    });
    if (res.error) throw new Error(res.error);
    jobId = res.job_id;
    if (poll) clearInterval(poll);
    poll = setInterval(pollStatus, 1000);
  } catch(e) {
    setSt("err","❌ " + e.message);
    byId("submitBtn").disabled = false;
  }
}

function resetForm() {
  ["f_title","f_link","f_note","f_ddl_date"].forEach(id => {
    const el = byId(id); if (el) el.value="";
  });
  updateDdlDateLabel();
  byId("f_ddl_hour").value = "";
  syncRequesterFromUser();
  byId("f_priority").value = "B";
  useCrowdin = null;
  inputMode = "";
  byId("togYes").classList.remove("on");
  byId("togNo").classList.remove("on");
  byId("f_link").readOnly = false;
  byId("uploadSection").style.display = "none";
  byId("linkRow").style.display = "none";
  removeAttachment(true);
  clearRichText();
  setInputMode("file");
  byId("progressCard").style.display = "none";
  byId("submitBtn").disabled = false;
}

window.addEventListener("load", () => {
  ["current_user","crowdin_token","apitable_api_key"].forEach(id => {
    const el = byId(id);
    if (el) el.addEventListener("input", () => updateConfigState(false));
    if (el) el.addEventListener("change", () => updateConfigState(false));
  });
  const task = byId("taskCard");
  if (task) task.addEventListener("click", e => {
    if (!isConfigReady()) {
      e.preventDefault();
      e.stopPropagation();
      requireConfig();
    }
  }, true);
  const drop = byId("drop");
  if (drop) {
    drop.addEventListener("click", e => {
      if (e.target && e.target.closest && e.target.closest("#removeFileBtn")) return;
      openFileChooser();
    });
    drop.addEventListener("keydown", e => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        openFileChooser();
      }
    });
  }
  updateConfigState(isConfigReady());
});
'''
    html_tmpl = html_tmpl.replace("</script>", extra_js + "\n</script>", 1)
    return html_tmpl

FINAL_HTML = build_html()

# ── HTTP Handler ──────────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): return

    def send_html(self, text: str, status: int=200) -> None:
        b = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type","text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers(); self.wfile.write(b)

    def send_json(self, data: Any, status: int=200) -> None:
        b = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type","application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers(); self.wfile.write(b)

    def read_json(self) -> Any:
        n = int(self.headers.get("Content-Length",0))
        return json.loads(self.rfile.read(n).decode("utf-8"))

    def do_GET(self) -> None:
        touch_client()
        p = urllib.parse.urlparse(self.path)
        if p.path in ("/","/index.html"):
            self.send_html(FINAL_HTML); return
        if p.path == "/api/state":
            self.send_json(STATE); return
        if p.path == "/api/status":
            jid = urllib.parse.parse_qs(p.query).get("job_id",[""])[0]
            job = JOBS.get(jid)
            if not job: self.send_json({"error":"not found"},404); return
            self.send_json({"status":job["status"],"log":job["log"],
                            "result":job.get("result"),"error":job.get("error")}); return
        self.send_json({"error":"not found"},404)

    def do_POST(self) -> None:
        p = urllib.parse.urlparse(self.path).path
        if p == "/api/heartbeat":
            touch_client()
            self.send_json({"ok": True}); return
        try: body = self.read_json()
        except Exception: self.send_json({"error":"bad body"},400); return
        touch_client()

        if p == "/api/state":
            next_state = dict(STATE)
            next_state.update(body)
            try:
                save_state(next_state)
            except Exception as exc:
                self.send_json({"error": f"保存本地配置失败: {exc}"}, 500); return
            STATE.clear()
            STATE.update(next_state)
            self.send_json({"ok":True}); return

        if p == "/api/ping":
            key = str(body.get("apitable_api_key","")).strip()
            self.send_json({"ok": aitable_ping(key)}); return

        if p == "/api/shutdown":
            self.send_json({"ok": True})
            threading.Thread(target=lambda: (time.sleep(0.3), os._exit(0)), daemon=True).start()
            return

        if p == "/api/submit":
            fb = body.get("file_bytes")
            if isinstance(fb, list): body["file_bytes"] = bytes(fb)
            else: body["file_bytes"] = b""
            # rich_html 直接透传
            if "rich_html" not in body: body["rich_html"] = ""

            jid = uuid.uuid4().hex
            JOBS[jid] = {"id":jid,"status":"pending","log":[],"result":None,"error":None,"payload":body}
            threading.Thread(target=run_submit_job, args=(JOBS[jid],), daemon=True).start()
            self.send_json({"job_id":jid}); return

        self.send_json({"error":"not found"},404)

# ── 入口 ──────────────────────────────────────────────────────────────────────
class _Srv(ThreadingHTTPServer):
    allow_reuse_address = True

def main() -> None:
    global STATE, LAST_CLIENT_SEEN
    STATE = load_state()
    port = choose_port(PORT)
    server = _Srv((HOST, port), Handler)
    LAST_CLIENT_SEEN = time.time()
    url = f"http://localhost:{port}"
    print(f"翻译任务提交工具  →  {url}")
    threading.Thread(target=client_watchdog, args=(server,), daemon=True).start()
    threading.Thread(target=lambda: (time.sleep(0.8), webbrowser.open(url)), daemon=True).start()
    try: server.serve_forever()
    except KeyboardInterrupt: print("\n已停止。")

if __name__ == "__main__":
    main()
