#!/usr/bin/env python3
"""
Crowdin Translation Checker
Paste Chinese + English draft → 3-column table with glossary terms → copy back to Crowdin.
Glossary auto-loaded from "Mafia War's Glossary.csv" in the same directory.
"""

import csv
import ast
import base64
import http.client
import http.server
import json
import math
import os
import re
import ssl
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request as _urllib_request
import webbrowser
import zipfile
import xml.etree.ElementTree as ET

PORT = 8765
CLIENT_STALE_SECONDS = 45
SHUTDOWN_GRACE_SECONDS = 1.5
# When packaged with PyInstaller (--onefile), __file__ points to a temporary
# extraction directory, not where the user actually placed the binary. Use the
# executable's directory in that case so the glossary CSV can sit next to the
# binary and be edited without rebuilding.
if getattr(sys, "frozen", False):
    SCRIPT_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

AICHAT_CREDS_FILE = os.path.join(SCRIPT_DIR, "aichat_creds.json")
_AI_SESSION_FILE = os.path.join(SCRIPT_DIR, "aichat_session.json")

# ── AIChat 内联实现 ────────────────────────────────────────────────────────────
HOST = "aichat.xinyoudi.com"
_ctx = ssl._create_unverified_context()

_RSA_PUBLIC_KEY = """-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAw9aTnFauSMuxMRKdIM6P
d8sfbvMUdtWewwGxbpwOhdETfpn2xE1AcAlNl53b/+EC+S3z7liaD+YnbbNbT+2w
I9k17Ey4nsi259ZU8WAi8064kkSAwSXQwBEX4tLPzTzD+VaK+f0q1+JwscaMqlOs
no6MwauirmcLdCDXszOeaIqLOdqo1JD9BTt2j6v74AEmxKLLm2G43lCaU5k6PWIC
RncHHPdqfadhvLC+hY2gS5aVWa+tsv7GCldLblnyFR6LnaYHNkQQ2DYoP2wwYZxH
k5t3g/gTuHZgH49qA3gaxYOL7kA+mxE/Xzku9FSA+P5hRYigiV7rHqZaFfiZB26o
QwIDAQAB
-----END PUBLIC KEY-----"""


def _http_get(host, path, cookies=None):
    conn = http.client.HTTPSConnection(host, context=_ctx, timeout=15)
    hdrs = {}
    if cookies:
        hdrs["Cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())
    conn.request("GET", path, headers=hdrs)
    resp = conn.getresponse()
    body = resp.read().decode("utf-8", errors="ignore")
    new_cookies = {}
    for h, v in resp.getheaders():
        if h.lower() == "set-cookie":
            k = v.split("=")[0].strip()
            vv = v.split("=")[1].split(";")[0]
            new_cookies[k] = vv
    location = next((v for h, v in resp.getheaders() if h.lower() == "location"), None)
    conn.close()
    return resp.status, new_cookies, location, body


def _http_post_json(host, path, payload, cookies=None):
    conn = http.client.HTTPSConnection(host, context=_ctx, timeout=15)
    body = json.dumps(payload).encode()
    hdrs = {"Content-Type": "application/json"}
    if cookies:
        hdrs["Cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())
    conn.request("POST", path, body=body, headers=hdrs)
    resp = conn.getresponse()
    result = resp.read().decode("utf-8", errors="ignore")
    new_cookies = {}
    for h, v in resp.getheaders():
        if h.lower() == "set-cookie":
            k = v.split("=")[0].strip()
            vv = v.split("=")[1].split(";")[0]
            new_cookies[k] = vv
    location = next((v for h, v in resp.getheaders() if h.lower() == "location"), None)
    conn.close()
    return resp.status, new_cookies, location, result


def _encrypt_password(password: str) -> str:
    try:
        from Crypto.PublicKey import RSA
        from Crypto.Cipher import PKCS1_v1_5
    except ImportError:
        raise RuntimeError("请安装 pycryptodome：pip install pycryptodome")
    payload = json.dumps({"password": password, "time": math.floor(time.time())})
    key = RSA.import_key(_RSA_PUBLIC_KEY)
    cipher = PKCS1_v1_5.new(key)
    return base64.b64encode(cipher.encrypt(payload.encode())).decode()


def _ai_load_creds():
    if not os.path.exists(AICHAT_CREDS_FILE):
        raise RuntimeError(f"找不到凭据文件 {AICHAT_CREDS_FILE}")
    with open(AICHAT_CREDS_FILE) as f:
        data = json.load(f)
    if "account" not in data or "password" not in data:
        raise RuntimeError(f"{AICHAT_CREDS_FILE} 缺少 account 或 password 字段")
    return data["account"], data["password"]


def _ai_save_session(aichat_session: str, xsrf_token: str):
    with open(_AI_SESSION_FILE, "w") as f:
        json.dump({"aichat_session": aichat_session, "xsrf_token": xsrf_token,
                   "saved_at": time.time()}, f)
    try:
        os.chmod(_AI_SESSION_FILE, 0o600)
    except Exception:
        pass


def _ai_load_session():
    if not os.path.exists(_AI_SESSION_FILE):
        return None, None
    try:
        with open(_AI_SESSION_FILE) as f:
            data = json.load(f)
        return data.get("aichat_session", ""), data.get("xsrf_token", "")
    except Exception:
        return None, None


def _ai_login(account: str = None, password: str = None):
    if account is None or password is None:
        account, password = _ai_load_creds()

    s, aichat_c, oauth_url, _ = _http_get(HOST, "/auth/redirect")
    if not oauth_url:
        raise RuntimeError("获取 OAuth 地址失败")

    parsed = urllib.parse.urlparse(oauth_url)
    oa_oauth_path = parsed.path + "?" + parsed.query

    s, oa_c, _, _ = _http_get("oa-core.xinyoudi.com", oa_oauth_path)

    enc_pwd = _encrypt_password(password)
    s, login_c, _, login_body = _http_post_json(
        "oa-core.xinyoudi.com",
        "/sfuser-api/sfhome/login",
        {"email": f"{account}@yottastudios.com", "password": enc_pwd},
        cookies=oa_c,
    )
    oa_c.update(login_c)
    login_data = json.loads(login_body)
    if login_data.get("error_code") not in ("0", 0):
        raise RuntimeError(f"OA 登录失败: {login_data.get('error_msg')}")

    s, _, callback_url, cb_body = _http_get("oa-core.xinyoudi.com", oa_oauth_path, cookies=oa_c)
    if not callback_url or "auth/callback" not in callback_url:
        # 服务端可能把 callback URL 放在 body 里（JSON 或 HTML/JS 重定向）
        for pattern in (
            r'"(?:redirect|callback|url|redirectUrl|redirect_url)"\s*:\s*"([^"]+auth/callback[^"]*)"',
            r"(?:window\.location|location\.href)\s*[=:]\s*['\"]([^'\"]+auth/callback[^'\"]*)['\"]",
            r'content=["\'][^"\']*url=([^"\']+auth/callback[^"\']*)["\']',
            r'(https?://[^\s"\'<>]+auth/callback[^\s"\'<>]*)',
        ):
            m = re.search(pattern, cb_body)
            if m:
                callback_url = m.group(1)
                break
    if not callback_url or "auth/callback" not in callback_url:
        raise RuntimeError(f"未拿到 callback URL: {callback_url}")

    parsed_cb = urllib.parse.urlparse(callback_url)
    cb_path = parsed_cb.path + "?" + parsed_cb.query
    s, final_c, _, _ = _http_get(HOST, cb_path, cookies=aichat_c)
    aichat_c.update(final_c)

    aichat_session = aichat_c.get("aichat_session", "")
    xsrf_token = aichat_c.get("XSRF-TOKEN", "")
    if not aichat_session:
        raise RuntimeError("登录后未获得 aichat_session")

    _ai_save_session(aichat_session, xsrf_token)
    return aichat_session, xsrf_token


def _get_valid_session():
    aichat_session, xsrf_token = _ai_load_session()
    if aichat_session:
        return aichat_session, xsrf_token
    return _ai_login()


def _do_ask(prompt, model_key, dialog_id, context_id, aichat_session, xsrf_token):
    body = json.dumps({
        "sys_lang": "zh-CN",
        "version_info": 1780282191327,
        "dialog_id": dialog_id,
        "context_id": context_id,
        "app_key": "chatgpt",
        "model_key": model_key,
        "content": prompt,
        "use_tools": [],
    }).encode()

    cookie_str = f"aichat_session={aichat_session}; XSRF-TOKEN={xsrf_token}"
    conn = http.client.HTTPSConnection(HOST, context=_ctx)
    conn.request("POST", "/web-api/openai/chat/ask", body=body, headers={
        "Content-Type": "application/json",
        "Cookie": cookie_str,
        "X-XSRF-TOKEN": xsrf_token,
    })
    resp = conn.getresponse()
    if resp.status != 200:
        raise RuntimeError(f"HTTP {resp.status}: {resp.read().decode()}")

    for h, v in resp.getheaders():
        if h.lower() == "set-cookie" and "aichat_session" in v:
            new_sess = v.split("=")[1].split(";")[0]
            _ai_save_session(new_sess, xsrf_token)
            break

    result_text = ""
    try:
        while True:
            line = resp.readline()
            if not line:
                break
            line = line.decode("utf-8", errors="ignore").strip()
            if not line.startswith("data:"):
                continue
            payload = line[5:].strip()
            if not payload:
                continue
            try:
                data = json.loads(payload)
            except json.JSONDecodeError:
                continue
            err = data.get("error_code")
            if err == "4500002":
                conn.close()
                return None
            if err not in (0, "0"):
                raise RuntimeError(f"API 错误: {data.get('error_msg')}")
            chunk = data.get("data", {})
            if chunk.get("type") == "message":
                result_text += chunk.get("text", "")
            elif chunk.get("type") == "chat" and chunk.get("is_finished"):
                break
    finally:
        conn.close()
    return result_text


def _aichat_ask(prompt, model_key, dialog_id, context_id):
    for _ in range(2):
        aichat_session, xsrf_token = _get_valid_session()
        result = _do_ask(prompt, model_key, dialog_id, context_id, aichat_session, xsrf_token)
        if result is not None:
            return result
        if os.path.exists(_AI_SESSION_FILE):
            os.remove(_AI_SESSION_FILE)
    raise RuntimeError("登录失败，请检查 aichat_creds.json 中的账密")
# ─────────────────────────────────────────────────────────────────────────────

GLOSSARY_FILE = "Mafia War's Glossary.csv"
AI_CONFIG_FILE = "crowdin_checker_ai_config.json"
REFERENCE_DIR = "reference"
DEFAULT_AI_CONFIG = {
    "model_key": "claude_opus_46",
    "translate_prompt": (
        "You are translating Mafia City localization content from Simplified Chinese to English.\n"
        "Return English only.\n"
        "Preserve every literal \\n exactly.\n"
        "Preserve inline color tags like [E7594C]...[-] exactly.\n"
        "Keep the original paragraph order.\n"
        "Follow the glossary when applicable.\n\n"
        "Glossary terms:\n{glossary_terms}\n\n"
        "Chinese source:\n{zh_text}"
    ),
    "check_prompt": (
        "You are reviewing an English localization draft against the Simplified Chinese source.\n"
        "Return the corrected English only.\n"
        "Preserve every literal \\n exactly as in the English draft.\n"
        "Preserve inline color tags like [E7594C]...[-] exactly.\n"
        "Keep the original paragraph order.\n"
        "Improve accuracy, fluency, consistency, and terminology.\n"
        "Follow the glossary when applicable.\n\n"
        "Glossary terms:\n{glossary_terms}\n\n"
        "Chinese source:\n{zh_text}\n\n"
        "Current English draft:\n{en_text}"
    ),
    "align_prompt": (
        "You are aligning an English localization draft to the Simplified Chinese source.\n"
        "Return English only.\n"
        "The English on the right must match the meaning of the Chinese on the left paragraph by paragraph.\n"
        "It does not need to match sentence by sentence, but it must not drift semantically.\n"
        "Preserve meaning, tone, literal \\n paragraph separators, and inline color tags exactly.\n"
        "Re-segment the English so each Chinese paragraph maps to the corresponding English paragraph.\n"
        "If a Chinese paragraph is a short title, label, or bracketed term, keep the English paragraph equally short instead of expanding it into a full explanation.\n"
        "Do not add notes or explanations.\n\n"
        "Glossary terms:\n{glossary_terms}\n\n"
        "Chinese source:\n{zh_text}\n\n"
        "Current English draft:\n{en_text}"
    ),
    "translate_templates": [
        {
            "id": "default",
            "name": "Default",
            "prompt": (
                "You are translating Mafia City localization content from Simplified Chinese to English.\n"
                "Return English only.\n"
                "Preserve every literal \\n exactly.\n"
                "Preserve inline color tags like [E7594C]...[-] exactly.\n"
                "Keep the original paragraph order.\n"
                "Follow the glossary when applicable.\n"
                "If reference content is provided, imitate its formatting, tone, and stylistic conventions where appropriate.\n\n"
                "Glossary terms:\n{glossary_terms}\n\n"
                "Reference content:\n{reference_text}\n\n"
                "Chinese source:\n{zh_text}"
            ),
            "reference_file": "",
        }
    ],
    "selected_translate_template_id": "default",
}
ALIGN_HARD_RULES = (
    "Alignment hard rules:\n"
    "1. The meaning on the English side must match the meaning on the Chinese side paragraph by paragraph.\n"
    "2. Do not force one-sentence-to-one-sentence matching. Semantic correspondence matters more than sentence count.\n"
    "3. Keep each Chinese paragraph aligned with the corresponding English paragraph; do not shift meaning into neighboring paragraphs.\n"
    "4. If a Chinese paragraph is a short title, label, bracketed term, or heading, the English output for that paragraph must also stay short and label-like. Do not expand it into an explanatory sentence.\n"
    "5. Do not add background explanation, lore, or extra context that is not present in the Chinese source.\n"
    "6. Preserve literal \\n separators and inline color tags exactly.\n"
    "7. Never move a heading/title paragraph to a neighboring index.\n"
    "8. Self-check before output: for each index i, ensure English paragraph i matches Chinese paragraph i semantically and is not shifted.\n"
    "9. Return English only.\n"
)


def load_aichat_creds():
    if not os.path.exists(AICHAT_CREDS_FILE):
        return {"configured": False, "account": ""}
    try:
        with open(AICHAT_CREDS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return {"configured": True, "account": data.get("account", "")}
    except Exception:
        return {"configured": False, "account": ""}


def save_aichat_creds(account, password):
    if not account or not password:
        raise ValueError("账号和密码不能为空")
    os.makedirs(SCRIPT_DIR, exist_ok=True)
    with open(AICHAT_CREDS_FILE, "w", encoding="utf-8") as f:
        json.dump({"account": account, "password": password}, f, ensure_ascii=False, indent=2)
    os.chmod(AICHAT_CREDS_FILE, 0o600)


def _glossary_candidates():
    """
    Return candidate glossary paths in priority order.
    Supports running as:
    - plain .py script
    - PyInstaller onedir executable
    - PyInstaller macOS .app (Finder launch)
    """
    candidates = []

    # 1) Next to the running executable/script
    candidates.append(os.path.join(SCRIPT_DIR, GLOSSARY_FILE))

    # 2) Current working directory (useful for terminal launches)
    candidates.append(os.path.join(os.getcwd(), GLOSSARY_FILE))

    # 3) If inside a macOS .app, also try the .app's parent folder
    #    .../MyApp.app/Contents/MacOS -> parent of .app bundle
    if ".app/Contents/MacOS" in SCRIPT_DIR:
        app_root = SCRIPT_DIR.split(".app/Contents/MacOS", 1)[0] + ".app"
        app_parent = os.path.dirname(app_root)
        candidates.append(os.path.join(app_parent, GLOSSARY_FILE))

    # de-duplicate while preserving order
    seen = set()
    ordered = []
    for p in candidates:
        norm = os.path.normpath(p)
        if norm not in seen:
            seen.add(norm)
            ordered.append(norm)
    return ordered


# Crowdin glossary export columns are fixed:
#   A (0)  -> Term [zh-CN]    源术语（中文）
#   M (12) -> Term [en-US]    目标术语（英文）
ZH_COL = 0
EN_COL = 12


def load_glossary():
    path = None
    for candidate in _glossary_candidates():
        if os.path.exists(candidate):
            path = candidate
            break
    if not path:
        return [], f"Glossary file not found: {GLOSSARY_FILE}"

    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))

    if len(rows) < 2:
        return [], "Glossary is empty or contains only a header row"

    terms = []
    for row in rows[1:]:
        zh = row[ZH_COL].strip() if len(row) > ZH_COL else ''
        en = row[EN_COL].strip() if len(row) > EN_COL else ''
        if zh and en:
            terms.append({'zh': zh, 'en': en})

    return terms, f"Loaded {len(terms)} glossary terms"


def _ai_config_path():
    return os.path.join(SCRIPT_DIR, AI_CONFIG_FILE)


def _default_ai_config():
    return json.loads(json.dumps(DEFAULT_AI_CONFIG))


def _reference_dir_path():
    return os.path.join(SCRIPT_DIR, REFERENCE_DIR)


def _normalize_translate_templates(value):
    defaults = _default_ai_config()["translate_templates"]
    raw = value if isinstance(value, list) and value else defaults
    out = []
    for idx, item in enumerate(raw):
        if not isinstance(item, dict):
            continue
        template_id = str(item.get("id") or f"template_{idx + 1}").strip()
        name = str(item.get("name") or f"Template {idx + 1}").strip()
        prompt = str(item.get("prompt") or defaults[0]["prompt"])
        reference_file = str(item.get("reference_file") or "").strip()
        out.append({
            "id": template_id,
            "name": name,
            "prompt": prompt,
            "reference_file": reference_file,
        })
    return out or defaults


def _load_docx_text(path):
    with zipfile.ZipFile(path) as zf:
        xml_bytes = zf.read("word/document.xml")
    root = ET.fromstring(xml_bytes)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = []
    for para in root.findall(".//w:p", ns):
        chunks = []
        for node in para.findall(".//w:t", ns):
            if node.text:
                chunks.append(node.text)
        text = "".join(chunks).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def list_reference_files():
    ref_dir = _reference_dir_path()
    if not os.path.isdir(ref_dir):
        return []
    names = []
    for name in sorted(os.listdir(ref_dir)):
        lower = name.lower()
        if lower.endswith(".docx") and not name.startswith("~") and not name.startswith("."):
            names.append(name)
    return names


def load_reference_text(filename):
    if not filename:
        return ""
    safe_name = os.path.basename(filename)
    path = os.path.join(_reference_dir_path(), safe_name)
    if not os.path.exists(path):
        raise ValueError(f"Reference file not found: {safe_name}")
    if not safe_name.lower().endswith(".docx"):
        raise ValueError("Only .docx reference files are supported")
    text = _load_docx_text(path)
    return text[:12000]


def load_ai_config():
    path = _ai_config_path()
    if not os.path.exists(path):
        return _default_ai_config()

    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (OSError, json.JSONDecodeError):
        return _default_ai_config()

    defaults = _default_ai_config()
    templates = _normalize_translate_templates(data.get("translate_templates"))
    selected_template_id = str(
        data.get("selected_translate_template_id") or defaults["selected_translate_template_id"]
    ).strip()
    if not any(t["id"] == selected_template_id for t in templates):
        selected_template_id = templates[0]["id"]

    return {
        "model_key": str(data.get("model_key") or defaults["model_key"]).strip(),
        "translate_prompt": str(data.get("translate_prompt") or defaults["translate_prompt"]),
        "check_prompt": str(data.get("check_prompt") or defaults["check_prompt"]),
        "align_prompt": str(data.get("align_prompt") or defaults["align_prompt"]),
        "translate_templates": templates,
        "selected_translate_template_id": selected_template_id,
    }


def save_ai_config(data):
    defaults = _default_ai_config()
    templates = _normalize_translate_templates(data.get("translate_templates"))
    selected_template_id = str(
        data.get("selected_translate_template_id") or defaults["selected_translate_template_id"]
    ).strip()
    if not any(t["id"] == selected_template_id for t in templates):
        selected_template_id = templates[0]["id"]

    config = {
        "model_key": str(data.get("model_key") or defaults["model_key"]).strip(),
        "translate_prompt": str(data.get("translate_prompt") or defaults["translate_prompt"]),
        "check_prompt": str(data.get("check_prompt") or defaults["check_prompt"]),
        "align_prompt": str(data.get("align_prompt") or defaults["align_prompt"]),
        "translate_templates": templates,
        "selected_translate_template_id": selected_template_id,
    }
    with open(_ai_config_path(), "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    return config


def normalize_ai_output(text: str) -> str:
    value = str(text or "").strip()

    # Drop leading explanation lines such as "Here is the corrected English:"
    lines = value.splitlines()
    while lines:
        head = lines[0].strip().lower()
        if (
            head.startswith("here is the corrected english")
            or head.startswith("corrected english")
            or head.startswith("revised english")
        ):
            lines.pop(0)
            continue
        break
    value = "\n".join(lines).strip()

    # Normalize common HTML color spans into the tool's native color tags.
    value = re.sub(
        r"<span\s+style=['\"]color\s*:\s*red\s*;?['\"]\s*>(.*?)</span>",
        r"[FF4D4F]\1[-]",
        value,
        flags=re.IGNORECASE | re.DOTALL,
    )
    value = re.sub(
        r"<span\s+style=['\"]color\s*:\s*green\s*;?['\"]\s*>(.*?)</span>",
        r"[56D364]\1[-]",
        value,
        flags=re.IGNORECASE | re.DOTALL,
    )
    value = re.sub(
        r"<span\s+style=['\"]color\s*:\s*#?ff0000\s*;?['\"]\s*>(.*?)</span>",
        r"[FF0000]\1[-]",
        value,
        flags=re.IGNORECASE | re.DOTALL,
    )
    value = re.sub(
        r"<span\s+style=['\"]color\s*:\s*#?00aa00\s*;?['\"]\s*>(.*?)</span>",
        r"[00AA00]\1[-]",
        value,
        flags=re.IGNORECASE | re.DOTALL,
    )
    value = re.sub(r"</?span[^>]*>", "", value, flags=re.IGNORECASE)

    return value.strip()


# 模型必须在对应品牌的对话下运行；dialog_id 是账号私有数据，不能写死，
# 否则换一个账号登录就会报错。统一做法：每次在当前账号下新建临时对话，用完即删。
_MODEL_BRANDS = (
    ("claude", "claude"),
    ("gemini", "gemini"),
    ("gpt", "openai"),
    ("deepseek", "deepseek"),
    ("grok", "grok"),
)


def _brand_for_model(model_key):
    for prefix, brand in _MODEL_BRANDS:
        if model_key.startswith(prefix):
            return brand
    return "openai"


def _dialog_api(path, body_dict):
    """POST 对话管理接口；session 过期（4500002）时自动重新登录重试一次。"""
    for attempt in range(2):
        session, xsrf = _get_valid_session()
        cookie = f"aichat_session={session}; XSRF-TOKEN={xsrf}"
        body = json.dumps({
            "sys_lang": "zh-CN", "version_info": 1780282191327,
            "app_key": "chatgpt", **body_dict,
        }).encode()
        req = _urllib_request.Request(
            f"https://{HOST}/web-api/{path}", data=body, method="POST",
        )
        req.add_header("Content-Type", "application/json")
        req.add_header("Cookie", cookie)
        req.add_header("X-XSRF-TOKEN", xsrf)
        with _urllib_request.urlopen(req, context=_ctx, timeout=15) as r:
            data = json.loads(r.read())
        err = data.get("error_code")
        if str(err) == "4500002" and attempt == 0:
            if os.path.exists(_AI_SESSION_FILE):
                os.remove(_AI_SESSION_FILE)
            continue
        if err not in (0, "0"):
            raise RuntimeError(f"AIChat 接口错误: {data.get('error_msg')}")
        return data
    raise RuntimeError("登录失败，请检查 aichat_creds.json 中的账密")


def _get_dialog_last_msg_id(dialog_id):
    """返回对话中最后一条消息的 id（用于识别新增消息）。"""
    data = _dialog_api("openai/chat/get-dialog-messages", {
        "dialog_id": dialog_id, "start_index": 0, "msg_num": 2,
    })
    msgs = data.get("data", {}).get("messages", [])
    return max((m.get("id", 0) for m in msgs), default=0)


def _get_new_assistant_text(dialog_id, after_id):
    """取 after_id 之后新增的 assistant 消息文本。"""
    data = _dialog_api("openai/chat/get-dialog-messages", {
        "dialog_id": dialog_id, "start_index": 0, "msg_num": 4,
    })
    msgs = data.get("data", {}).get("messages", [])
    for msg in reversed(msgs):
        if msg.get("id", 0) > after_id and msg.get("role") == "assistant":
            return "".join(
                c.get("text", "") for c in msg.get("content", []) if isinstance(c, dict)
            )
    return ""


def _create_tmp_dialog(model_key):
    """在当前账号下新建对应品牌的临时对话，返回 (dialog_id, context_id)。"""
    data = _dialog_api("openai/chat/add-dialog", {
        "brand": _brand_for_model(model_key),
        "model_key": model_key, "title": "翻译检查临时",
    })
    d = data.get("data", {})
    return d["id"], d["last_context_id"]


def _delete_dialog(dialog_id):
    """删除指定对话（用于清理临时 dialog）。"""
    try:
        _dialog_api("openai/chat/delete-dialog", {"id": dialog_id})
    except Exception:
        pass  # 删除失败不阻断主流程


def call_ai_chat(messages, config, temperature=0.2, max_tokens_override=None):
    model_key = str(config.get("model_key") or "claude_opus_46").strip()
    if not model_key:
        raise ValueError("Please select a Model first")
    prompt = "\n".join(
        str(m.get("content") or "")
        for m in messages
        if m.get("role") == "user"
    )

    # 所有模型统一：在当前账号下新建临时 dialog（无历史干扰、不依赖写死的
    # dialog_id），用完即删。Claude 不走 SSE 返回文本，需从对话历史取增量。
    tmp_dialog_id, tmp_context_id = _create_tmp_dialog(model_key)
    try:
        if model_key.startswith("claude"):
            prev_id = _get_dialog_last_msg_id(tmp_dialog_id)
            _aichat_ask(prompt, model_key=model_key,
                        dialog_id=tmp_dialog_id, context_id=tmp_context_id)
            raw_text = _get_new_assistant_text(tmp_dialog_id, prev_id)
        else:
            raw_text = _aichat_ask(prompt, model_key=model_key,
                                   dialog_id=tmp_dialog_id,
                                   context_id=tmp_context_id)
    finally:
        _delete_dialog(tmp_dialog_id)

    return {
        "text": normalize_ai_output(raw_text),
        "raw_text": raw_text,
        "model": model_key,
    }


def build_glossary_prompt(terms, source_text):
    hits = []
    seen = set()
    for term in terms:
        zh = term.get("zh", "")
        en = term.get("en", "")
        if zh and en and zh in source_text and zh not in seen:
            hits.append(f"{zh} = {en}")
            seen.add(zh)
    return "\n".join(hits[:200])


def parse_raw_with_seps(raw_text):
    tokens = str(raw_text or "").split("\\n")
    paras = []
    seps = []
    pending_empty = 0
    for tok in tokens:
        trimmed = tok.strip()
        if trimmed:
            if paras:
                seps.append("\\n" * (1 + pending_empty))
            paras.append(trimmed)
            pending_empty = 0
        else:
            pending_empty += 1
    return {"paras": paras, "seps": seps}


def build_align_json_prompt(zh_text, en_text, glossary_prompt, custom_prompt=""):
    zh = parse_raw_with_seps(zh_text)
    en = parse_raw_with_seps(en_text)
    zh_lines = [
        f"{idx + 1}. {para}"
        for idx, para in enumerate(zh["paras"])
    ]
    en_lines = [
        f"{idx + 1}. {para}"
        for idx, para in enumerate(en["paras"])
    ]
    extra = ""
    safe_custom = str(custom_prompt or "").strip()
    if safe_custom:
        extra = "\n\nAdditional alignment preference (must not override the hard rules above):\n" + safe_custom
    return (
        ALIGN_HARD_RULES
        + "\nReturn valid JSON only. No markdown fences. No explanation.\n"
        + "Your entire reply must start with { or [ and end with } or ].\n"
        + "JSON schema:\n"
        + '{\"paragraphs\": [\"english paragraph 1\", \"english paragraph 2\"]}\n'
        + "The paragraphs array length must exactly equal the number of Chinese paragraphs.\n\n"
        + "Glossary terms:\n"
        + (glossary_prompt or "(no matched glossary terms)")
        + "\n\nChinese paragraphs:\n"
        + ("\n".join(zh_lines) if zh_lines else "(empty)")
        + "\n\nCurrent English draft:\n"
        + ("\n".join(en_lines) if en_lines else "(empty)")
        + extra
    )


def build_align_repair_prompt(zh_text, en_text, glossary_prompt, broken_output, expected_count):
    zh = parse_raw_with_seps(zh_text)
    en = parse_raw_with_seps(en_text)
    zh_lines = [f"{idx + 1}. {para}" for idx, para in enumerate(zh["paras"])]
    en_lines = [f"{idx + 1}. {para}" for idx, para in enumerate(en["paras"])]
    return (
        ALIGN_HARD_RULES
        + "\nYour previous output failed validation.\n"
        + f"You MUST return exactly {expected_count} paragraphs in JSON array `paragraphs`.\n"
        + "Do not merge or drop any paragraph index.\n"
        + "Return valid JSON only. No markdown fences. No explanation.\n"
        + "JSON schema:\n"
        + '{\"paragraphs\": [\"english paragraph 1\", \"english paragraph 2\"]}\n\n'
        + "Glossary terms:\n"
        + (glossary_prompt or "(no matched glossary terms)")
        + "\n\nChinese paragraphs:\n"
        + ("\n".join(zh_lines) if zh_lines else "(empty)")
        + "\n\nCurrent English draft:\n"
        + ("\n".join(en_lines) if en_lines else "(empty)")
        + "\n\nYour previous invalid output (for correction only):\n"
        + str(broken_output or "")
    )


def _extract_json_blob(text):
    value = str(text or "").strip()
    fenced = re.search(r"```(?:json)?\s*([\s\S]*?)```", value, flags=re.IGNORECASE)
    if fenced:
        return fenced.group(1).strip()
    obj_start = value.find("{")
    obj_end = value.rfind("}")
    if obj_start != -1 and obj_end != -1 and obj_end > obj_start:
        return value[obj_start:obj_end + 1]
    arr_start = value.find("[")
    arr_end = value.rfind("]")
    if arr_start != -1 and arr_end != -1 and arr_end > arr_start:
        return value[arr_start:arr_end + 1]
    raise ValueError("AI Align did not return JSON")


def _try_parse_align_fallback(text, expected_count, fallback_seps):
    raw = normalize_ai_output(text)
    if not raw:
        return None

    # If the model already returned content with the correct paragraph count,
    # accept it even when it ignored the JSON wrapper requirement.
    parsed = parse_raw_with_seps(raw)
    if len(parsed["paras"]) == expected_count:
        return raw

    # Common fallback: numbered lines like "1. ...", "2. ...".
    lines = [line.strip() for line in raw.splitlines() if line.strip()]
    numbered = []
    for line in lines:
        m = re.match(r"^\d+\.\s*(.*)$", line)
        if m:
            numbered.append(normalize_ai_output(m.group(1)))
    if len(numbered) == expected_count:
        result = ""
        for idx, para in enumerate(numbered):
            result += para
            if idx < len(numbered) - 1:
                result += fallback_seps[idx] if idx < len(fallback_seps) else "\\n"
        return result

    # Another common fallback: one paragraph per physical line.
    if len(lines) == expected_count:
        result = ""
        for idx, para in enumerate(lines):
            result += normalize_ai_output(para)
            if idx < len(lines) - 1:
                result += fallback_seps[idx] if idx < len(fallback_seps) else "\\n"
        return result

    return None


def parse_align_json_response(text, expected_count, fallback_seps):
    try:
        blob = _extract_json_blob(text)
    except ValueError:
        fallback = _try_parse_align_fallback(text, expected_count, fallback_seps)
        if fallback is not None:
            return fallback
        raise
    try:
        data = json.loads(blob)
    except json.JSONDecodeError as e:
        try:
            data = ast.literal_eval(blob)
        except (ValueError, SyntaxError):
            fallback = _try_parse_align_fallback(text, expected_count, fallback_seps)
            if fallback is not None:
                return fallback
            raise ValueError("AI Align returned invalid JSON") from e

    if isinstance(data, list):
        paras = data
    else:
        paras = data.get("paragraphs")
    if not isinstance(paras, list):
        fallback = _try_parse_align_fallback(text, expected_count, fallback_seps)
        if fallback is not None:
            return fallback
        raise ValueError("AI Align JSON must contain a 'paragraphs' array")
    cleaned = [normalize_ai_output(str(item or "").strip()) for item in paras]
    if len(cleaned) != expected_count:
        fallback = _try_parse_align_fallback(text, expected_count, fallback_seps)
        if fallback is not None:
            return fallback
        raise ValueError(
            f"AI Align returned {len(cleaned)} paragraphs, expected {expected_count}"
        )

    result = ""
    for idx, para in enumerate(cleaned):
        result += para
        if idx < len(cleaned) - 1:
            result += fallback_seps[idx] if idx < len(fallback_seps) else "\\n"
    return result


def _is_zh_heading_like(text):
    value = str(text or "").strip()
    if not value:
        return False
    if re.match(r"^[\[\【].{1,40}[\]\】]$", value):
        return True
    if len(value) <= 24 and not re.search(r"[。！？；，,.!?;:]", value):
        return True
    return False


def _is_en_heading_like(text):
    value = str(text or "").strip()
    if not value:
        return False
    if re.match(r"^\[[^\]]{1,60}\]$", value):
        return True
    words = value.split()
    if len(words) <= 8 and not re.search(r"[.!?]", value):
        return True
    return False


def _fix_adjacent_heading_shift(zh_paras, en_paras):
    if len(zh_paras) != len(en_paras):
        return en_paras
    fixed = list(en_paras)
    changed = False
    for i in range(len(fixed) - 1):
        zh_i_heading = _is_zh_heading_like(zh_paras[i])
        zh_j_heading = _is_zh_heading_like(zh_paras[i + 1])
        en_i_heading = _is_en_heading_like(fixed[i])
        en_j_heading = _is_en_heading_like(fixed[i + 1])
        # If heading/non-heading pattern is inverted across adjacent indexes,
        # swap once to recover from the common +1 shift.
        if zh_i_heading == en_j_heading and zh_j_heading == en_i_heading and (en_i_heading != en_j_heading):
            fixed[i], fixed[i + 1] = fixed[i + 1], fixed[i]
            changed = True
    return fixed if changed else en_paras


def _looks_like_truncated_align_json(text):
    value = str(text or "").strip()
    if not value:
        return False

    starts_json_like = value.startswith("{") or value.startswith("[") or "```json" in value.lower()
    if not starts_json_like:
        return False

    # Common truncated-response signal: unbalanced structure.
    # Simple counting is enough here and keeps Python 3.9 compatibility.
    return value.count("{") != value.count("}") or value.count("[") != value.count("]")


def fill_prompt_template(template, zh_text="", en_text="", glossary_prompt="", reference_text=""):
    safe_template = str(template or "").strip()
    if not safe_template:
        raise ValueError("Prompt cannot be empty")
    try:
        return safe_template.format(
            zh_text=zh_text,
            en_text=en_text,
            glossary_terms=glossary_prompt or "(no matched glossary terms)",
            reference_text=reference_text or "(no reference content)",
        )
    except KeyError as e:
        raise ValueError(f"Unsupported prompt placeholder: {e.args[0]}") from e


def build_single_prompt_messages(prompt_text):
    return [
        {"role": "user", "content": prompt_text},
    ]


# ─────────────────────────────────────────────────────────────────────────────
# Embedded frontend
# ─────────────────────────────────────────────────────────────────────────────
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Crowdin Translation Checker</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;
     background:#0d1117;color:#c9d1d9;min-height:100vh}

/* ── Header ── */
.hdr{background:#161b22;border-bottom:1px solid #30363d;padding:12px 24px;
     display:flex;align-items:center;gap:16px;position:sticky;top:0;z-index:100}
.hdr h1{font-size:15px;font-weight:600;color:#f0f6fc;white-space:nowrap}
.gloss-tag{margin-left:auto;font-size:12px;padding:4px 12px;border-radius:20px;
           background:#122d20;color:#56d364;white-space:nowrap}
.gloss-tag.warn{background:#32100b;color:#f85149}
.gloss-tag.loading{background:#1c2a3a;color:#8b949e}

/* ── Buttons ── */
.btn{padding:7px 16px;border-radius:6px;border:none;cursor:pointer;
     font-size:13px;font-weight:500;transition:background .15s;white-space:nowrap}
.btn:disabled{opacity:.4;cursor:not-allowed}
.btn-blue{background:#1f6feb;color:#fff}.btn-blue:hover:not(:disabled){background:#388bfd}
.btn-green{background:#238636;color:#fff}.btn-green:hover:not(:disabled){background:#2ea043}
.btn-sky{background:#0969da;color:#fff}.btn-sky:hover:not(:disabled){background:#1f6feb}
.btn-gray{background:#21262d;color:#c9d1d9;border:1px solid #30363d}
.btn-gray:hover:not(:disabled){background:#30363d}
.btn-sm{padding:4px 10px;font-size:11px;border-radius:5px;border:none;cursor:pointer;
        font-weight:500;transition:background .15s;white-space:nowrap;
        background:#21262d;color:#8b949e;border:1px solid #30363d}
.btn-sm:hover{background:#30363d;color:#c9d1d9}
.btn-sm.active{background:#1c2a3a;color:#79c0ff;border-color:#1f4068}

/* ── Input panel ── */
.panel{padding:16px 24px;background:#161b22;border-bottom:1px solid #30363d}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:12px}
.field label{display:block;font-size:11px;color:#8b949e;margin-bottom:6px;
             font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.field label span{text-transform:none;font-weight:400;color:#484f58}
textarea{width:100%;background:#0d1117;border:1px solid #30363d;border-radius:6px;
         padding:9px 11px;color:#c9d1d9;font-size:12.5px;
         font-family:'SFMono-Regular',Consolas,monospace;resize:vertical;
         min-height:130px;line-height:1.6;overflow-y:auto;overflow-x:hidden}
select{width:100%;background:#0d1117;border:1px solid #30363d;border-radius:6px;
       padding:9px 11px;color:#c9d1d9;font-size:12.5px;min-height:40px}
textarea:focus,select:focus{outline:none;border-color:#388bfd}
.mini-grid{display:grid;grid-template-columns:1fr auto;gap:10px;margin-bottom:12px}
.creds-grid{display:grid;grid-template-columns:1fr 1fr auto;gap:10px;margin-bottom:12px}
.creds-tag{font-size:11px;font-weight:400;padding:2px 8px;border-radius:10px;margin-left:6px;
           background:#122d20;color:#56d364}
.creds-tag.warn{background:#32100b;color:#f85149}
.mini-field input{width:100%;background:#0d1117;border:1px solid #30363d;border-radius:6px;
                  padding:9px 11px;color:#c9d1d9;font-size:12.5px}
.mini-field input:focus{outline:none;border-color:#388bfd}
.mini-field label{display:block;font-size:11px;color:#8b949e;margin-bottom:6px;
                  font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.ai-actions{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}
.prompt-grid{display:grid;grid-template-columns:1.1fr 0.9fr 1fr auto;gap:10px;margin-bottom:12px}
.prompt-field textarea{min-height:120px;font-size:12px}
.template-toolbar{display:flex;gap:8px;align-items:end;flex-wrap:wrap;margin-bottom:12px}
.template-toolbar .mini-field{min-width:180px;flex:1}
.template-actions{display:flex;gap:8px;align-items:end;flex-wrap:wrap}
.usage-box{margin-top:10px;padding:10px 12px;border:1px solid #30363d;border-radius:6px;
           background:#0d1117;font-size:12px;color:#8b949e;display:none;line-height:1.6}
.usage-box.on{display:block}
.debug-box{margin-top:10px;padding:12px;border:1px solid #30363d;border-radius:6px;background:#0d1117;display:none}
.debug-box.on{display:block}
.debug-hdr{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px}
.debug-hdr strong{font-size:12px;color:#f0f6fc;letter-spacing:.04em}
.debug-box textarea{min-height:130px;font-size:12px}
.row-btns{display:flex;gap:10px;align-items:center}
.status{font-size:13px;padding:5px 12px;border-radius:6px;display:none}
.status.info{background:#1c3557;color:#79c0ff;display:inline-block}
.status.ok  {background:#122d20;color:#56d364;display:inline-block}
.status.err {background:#32100b;color:#f85149;display:inline-block}

/* ── Blocks wrapper ── */
.blocks-wrap{padding:16px 24px}
.blocks-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.blocks-hdr>span{font-size:13px;color:#8b949e}
.blocks-actions{display:flex;gap:8px}

/* ── Paragraph block ── */
.para-block{background:#161b22;border:1px solid #30363d;border-radius:8px;
            margin-bottom:12px;overflow:hidden}
.block-hdr{display:flex;align-items:center;gap:10px;padding:8px 14px;
           background:#1c2128;border-bottom:1px solid #30363d;flex-wrap:wrap}
.block-label{font-size:12px;font-weight:600;color:#8b949e;
             background:#21262d;padding:2px 9px;border-radius:10px;white-space:nowrap}
.block-pills{flex:1;display:flex;flex-wrap:wrap;gap:4px}
.pill{display:inline-block;background:#1c2a3a;color:#79c0ff;padding:2px 9px;
      border-radius:12px;font-size:11px;white-space:nowrap;border:1px solid #1f4068}
.pill .pill-zh{color:#c9d1d9}
.pill .pill-eq{color:#484f58;margin:0 4px}
.no-term{color:#484f58;font-size:11px}
.term-hit{text-decoration:underline dashed #79c0ff;
          text-underline-offset:3px;cursor:help}

/* ── Help / tutorial panel ── */
.help-panel{background:#0d2030;border-bottom:1px solid #1f4068;padding:14px 24px;
            font-size:12.5px;line-height:1.7;color:#8b949e;display:none}
.help-panel.on{display:block}
.help-panel h3{font-size:12px;color:#79c0ff;text-transform:uppercase;
               letter-spacing:.6px;font-weight:600;margin-bottom:8px}
.help-panel ol{margin-left:20px;color:#b0b8c0}
.help-panel ol li{margin-bottom:4px}
.help-panel code{background:#161b22;padding:1px 6px;border-radius:4px;
                 font-family:'SFMono-Regular',Consolas,monospace;font-size:11.5px;
                 color:#79c0ff}
.help-toggle{background:#21262d;color:#8b949e;border:1px solid #30363d;
             border-radius:50%;width:26px;height:26px;cursor:pointer;
             font-size:13px;font-weight:600;line-height:1;padding:0}
.help-toggle:hover{background:#30363d;color:#c9d1d9}
.help-toggle.active{background:#1c2a3a;color:#79c0ff;border-color:#1f4068}

/* ── Merge / unmerge buttons ── */
.btn-merge{padding:4px 10px;font-size:11px;border-radius:5px;border:1px solid #1f4068;
           background:#0d2030;color:#79c0ff;cursor:pointer;font-weight:500;
           white-space:nowrap;transition:background .15s}
.btn-merge:hover:not(:disabled){background:#1c2a3a}
.btn-merge:disabled{opacity:.3;cursor:not-allowed}

/* ── Paragraph divider inside a merged group ── */
.para-divider{padding:5px 14px;color:#79c0ff;font-size:10.5px;
              background:#0d2030;border-top:1px solid #1f4068;
              border-bottom:1px solid #1f4068;letter-spacing:.4px;
              font-weight:500}
.para-block.is-merged{border-color:#1f4068}
.para-block.is-merged .block-hdr{background:#0d2030}

/* ── Sentence rows inside a block ── */
.block-body{width:100%}
.sent-row{display:grid;grid-template-columns:38px 1fr 1fr;border-bottom:1px solid #21262d}
.sent-row:last-child{border-bottom:none}
.sent-row:hover{background:#1c2128}
.sent-num{display:flex;align-items:flex-start;justify-content:center;
          padding:10px 4px;color:#484f58;font-size:11px;font-weight:600;
          border-right:1px solid #21262d;min-width:38px}
.sent-zh{padding:10px 12px;font-size:13px;line-height:1.8;color:#c9d1d9;
         border-right:1px solid #21262d;word-break:break-all}
.sent-en{padding:8px 10px}
.sent-en textarea{width:100%;background:transparent;border:1px solid transparent;
                  border-radius:4px;padding:4px 8px;color:#c9d1d9;font-size:13px;
                  font-family:'SFMono-Regular',Consolas,monospace;resize:none;
                  line-height:1.8;overflow:hidden;min-height:32px;display:block}
.sent-en textarea:hover{border-color:#30363d;background:#0d1117}
.sent-en textarea:focus{border-color:#388bfd;background:#0d1117;outline:none}
.en-preview{margin-top:6px;padding:6px 8px;border:1px dashed #30363d;border-radius:4px;
            background:#11161d;color:#8b949e;font-size:12px;line-height:1.7;word-break:break-word}
.en-preview-line{margin:0;cursor:pointer;border-radius:4px;padding:4px 6px}
.en-preview-line:hover{background:#1a222d}
.en-preview-line + .en-preview-line{margin-top:6px;padding-top:6px;border-top:1px dashed #30363d}
.diff-new{color:#ff6b6b;font-weight:600}
.diff-old{color:#56d364;font-weight:600}

/* ── Combined view (per block) ── */
.combined-view{display:none;padding:0}
.combined-body{display:grid;grid-template-columns:1fr 1fr}
.combined-zh{padding:12px 14px;font-size:13px;line-height:1.9;color:#c9d1d9;
             border-right:1px solid #21262d;word-break:break-all}
.sent-sep{display:block;color:#484f58;font-size:11px;margin:4px 0 2px}
.combined-en{padding:10px 12px}
.combined-en textarea{width:100%;background:#0d1117;border:1px solid #30363d;
                      border-radius:6px;padding:10px 12px;color:#c9d1d9;
                      font-size:13px;font-family:'SFMono-Regular',Consolas,monospace;
                      resize:none;line-height:1.9;overflow:hidden;display:block}
.combined-en textarea:focus{border-color:#388bfd;outline:none}
.combined-en .en-preview{margin-top:10px}

/* ── Preview modal ── */
.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.82);
         z-index:200;overflow-y:auto;padding:24px 16px}
.overlay.on{display:flex;justify-content:center;align-items:flex-start}
.modal{background:#161b22;border:1px solid #30363d;border-radius:12px;
       width:100%;max-width:980px;margin:auto}
.modal-hdr{padding:14px 20px;border-bottom:1px solid #30363d;
           display:flex;align-items:center;justify-content:space-between}
.modal-hdr h2{font-size:15px;font-weight:600;color:#f0f6fc}
.close-btn{background:none;border:none;color:#8b949e;font-size:22px;
           cursor:pointer;line-height:1;padding:0 4px}
.close-btn:hover{color:#c9d1d9}
.modal-body{padding:20px}
.preview-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
.preview-col h3{font-size:11px;color:#8b949e;text-transform:uppercase;
                letter-spacing:.5px;margin-bottom:10px;font-weight:600}
.preview-box{background:#0d1117;border:1px solid #30363d;border-radius:8px;
             padding:14px 16px;font-size:13px;line-height:1.9;
             max-height:62vh;overflow-y:auto;word-break:break-word}
.pv-para{margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid #21262d}
.pv-para:last-child{border-bottom:none;margin-bottom:0}
.modal-ftr{padding:14px 20px;border-top:1px solid #30363d;
           display:flex;gap:10px;justify-content:flex-end}

/* ── Toast ── */
.toast{position:fixed;bottom:22px;right:22px;background:#238636;color:#fff;
       padding:10px 20px;border-radius:8px;font-size:13px;z-index:999;
       display:none;box-shadow:0 4px 14px rgba(0,0,0,.5)}
.toast.on{display:block;animation:slideUp .2s ease}
@keyframes slideUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}
</style>
</head>
<body>

<!-- ── Header ── -->
<div class="hdr">
  <h1>🎮 Crowdin Translation Checker</h1>
  <button class="help-toggle" id="helpToggle" onclick="toggleHelp()" title="Show / hide instructions">?</button>
  <div class="gloss-tag loading" id="glossTag">Loading glossary…</div>
</div>

<!-- ── Help / tutorial panel ── -->
<div class="help-panel" id="helpPanel">
  <h3>How to use</h3>
  <ol>
    <li>Paste the <strong>Chinese source</strong> on the left. The <strong>English draft</strong> on the right is <em>optional</em> &mdash; paste one to review/edit, or leave it empty to translate from scratch directly in the table below.</li>
    <li>Use the literal token <code>\n</code> to separate paragraphs &mdash; in both inputs when present, ideally one-to-one. If the English side has fewer (or no) <code>\n</code>, the output will follow the Chinese source's <code>\n</code> structure.</li>
    <li>Click <strong>Build Table</strong>. Each paragraph becomes a row that you can edit sentence by sentence, or toggle <strong>Merge view</strong> to edit a whole paragraph at once.</li>
    <li>Glossary hits are shown as <code>中文 = English</code> pills above each block; matched substrings in the source get a <span class="term-hit">dashed underline</span> (hover for the English term).</li>
    <li><strong>⬇ Merge next</strong> visually combines two adjacent blocks so you can review related paragraphs side by side. <em>This is view-only</em> &mdash; the copied output is unaffected. Press <strong>↑ Unmerge</strong> on a merged block to split it back apart.</li>
    <li><strong>Preview</strong> renders edits live: color codes like <code>[E7594C]…[-]</code> become colored text, and any literal <code>\n</code> you type inside an edit area becomes a real line break. Both are preserved verbatim in the copied output.</li>
    <li>Click <strong>Copy</strong> &mdash; the result always preserves the <em>exact</em> <code>\n</code> structure of your English draft (e.g. <code>\n\n</code> stays <code>\n\n</code>), no matter how you merge / unmerge. Nothing is auto-added.</li>
  </ol>
</div>

<!-- ── Input panel ── -->
<div class="panel">
  <div class="mini-grid">
    <div class="mini-field">
      <label>Model</label>
      <select id="aiModelKey" onchange="saveAiConfig({silent:true})">
        <optgroup label="── 最高质量 ──">
          <option value="claude_opus_46">Claude Opus</option>
          <option value="claude_sonnet_46">Claude Sonnet</option>
          <option value="gpt_5">GPT-5</option>
        </optgroup>
        <optgroup label="── 深度推理 ──">
          <option value="deepseek_reasoner">DeepSeek 推理链</option>
        </optgroup>
        <optgroup label="── 均衡 ──">
          <option value="gemini_31_flash_image">Gemini Flash（默认）</option>
          <option value="gpt_4o">GPT-4o</option>
          <option value="gpt_41">GPT-4.1</option>
          <option value="deepseek_v3_1">DeepSeek V3</option>
          <option value="gemini_15_pro">Gemini Pro</option>
          <option value="grok_3">Grok 3</option>
        </optgroup>
        <optgroup label="── 快速 ──">
          <option value="gpt_4o_mini">GPT-4o Mini</option>
          <option value="deepseek_chat">DeepSeek 快速版</option>
        </optgroup>
      </select>
    </div>
    <div class="ai-actions">
      <button class="btn btn-gray" onclick="saveAiConfig()">💾 Save</button>
    </div>
  </div>
  <div class="creds-grid">
    <div class="mini-field">
      <label>OA 账号 <span id="credsTag" class="creds-tag warn">未配置</span></label>
      <input id="credsAccount" type="text" placeholder="邮箱 @ 前的用户名，如 zhangsan">
    </div>
    <div class="mini-field">
      <label>密码</label>
      <input id="credsPassword" type="password" placeholder="OA 密码">
    </div>
    <div class="ai-actions">
      <button class="btn btn-gray" onclick="saveCreds()">🔑 保存账密</button>
    </div>
  </div>
  <div class="template-toolbar">
    <div class="mini-field">
      <label>Translate Template</label>
      <select id="translateTemplateSelect" onchange="onTemplateChange()"></select>
    </div>
    <div class="mini-field">
      <label>Reference</label>
      <select id="referenceSelect" onchange="onReferenceChange()"></select>
    </div>
    <div class="mini-field">
      <label>Template Name</label>
      <input id="templateName" type="text" placeholder="Template name">
    </div>
    <div class="template-actions">
      <button class="btn btn-gray" onclick="newTemplate()">New Template</button>
      <button class="btn btn-gray" onclick="deleteTemplate()">Delete Template</button>
    </div>
  </div>
  <div class="prompt-grid">
    <div class="prompt-field">
      <label>Translate Prompt</label>
      <textarea id="translatePrompt"
        placeholder="Use {zh_text} / {glossary_terms} / {reference_text} placeholders"></textarea>
    </div>
    <div class="prompt-field">
      <label>Check Prompt</label>
      <textarea id="checkPrompt"
        placeholder="Use {zh_text} / {en_text} / {glossary_terms}"></textarea>
    </div>
    <div class="prompt-field">
      <label>Align Prompt</label>
      <textarea id="alignPrompt"
        placeholder="Use {zh_text} / {en_text} / {glossary_terms}"></textarea>
    </div>
    <div class="ai-actions">
      <button class="btn btn-gray" onclick="saveAiConfig()">💾 Save Prompts</button>
    </div>
  </div>
  <div class="grid2">
    <div class="field">
      <label>Chinese Source <span>(use \n to separate paragraphs; supports [RRGGBB]…[-] color codes)</span></label>
      <textarea id="zh"
        placeholder="Paste Chinese source, e.g.&#10;各位市民即将可以选择离开本城市。但为了维护本城秩序，请遵守规定。\n[E7594C]请注意！[-]活动期间请保持冷静。"></textarea>
    </div>
    <div class="field">
      <label>English Draft <span>(optional — leave empty to translate from scratch; use \n to separate paragraphs, should match the Chinese 1:1)</span></label>
      <textarea id="en"
        placeholder="Optional. Paste an existing English draft to review, or leave empty and fill it in below.&#10;e.g. Citizens will soon be able to leave the city.\n[E7594C]Please note![-] Stay calm during the event."></textarea>
    </div>
  </div>
  <div class="row-btns">
    <button class="btn btn-blue" onclick="buildBlocks()">📊 Build Table</button>
    <button class="btn btn-sky" onclick="runAiAlign()">🧩 AI Align</button>
    <button class="btn btn-sky" onclick="runAiTranslate()">🤖 AI Translate</button>
    <button class="btn btn-gray" onclick="runAiCheck()">🩺 AI Check</button>
    <div class="status" id="status"></div>
  </div>
  <div class="usage-box" id="usageBox"></div>
  <div class="debug-box" id="debugBox">
    <div class="debug-hdr">
      <strong>AI RAW OUTPUT</strong>
      <button class="btn btn-gray" onclick="copyRawOutput()">Copy Raw Output</button>
    </div>
    <textarea id="debugRawOutput" readonly placeholder="The model's raw response will appear here when available."></textarea>
  </div>
</div>

<!-- ── Paragraph blocks ── -->
<div class="blocks-wrap" id="blocksWrap" style="display:none">
  <div class="blocks-hdr">
    <span><strong id="cnt" style="color:#f0f6fc">0</strong> paragraph(s)</span>
    <div class="blocks-actions">
      <button class="btn btn-sky" onclick="showPreview()">👁 Preview</button>
      <button class="btn btn-green" onclick="doCopy()">📋 Copy Final Result</button>
    </div>
  </div>
  <div id="blocks"></div>
</div>

<!-- ── Preview modal ── -->
<div class="overlay" id="overlay">
  <div class="modal">
    <div class="modal-hdr">
      <h2>Preview &mdash; Colors &amp; Paragraph Layout</h2>
      <button class="close-btn" onclick="closePreview()">×</button>
    </div>
    <div class="modal-body">
      <div class="preview-grid">
        <div class="preview-col">
          <h3>Chinese Source</h3>
          <div class="preview-box" id="pvZh"></div>
        </div>
        <div class="preview-col">
          <h3>English (current edit)</h3>
          <div class="preview-box" id="pvEn"></div>
        </div>
      </div>
    </div>
    <div class="modal-ftr">
      <button class="btn btn-gray" onclick="closePreview()">Close</button>
      <button class="btn btn-green" onclick="confirmAndCopy()">✓ Confirm &amp; Copy to Clipboard</button>
    </div>
  </div>
</div>

<!-- ── Toast ── -->
<div class="toast" id="toast">✓ Copied! Paste back into Crowdin.</div>

<script>
// ── State ──────────────────────────────────────────────────────────────────
let glossary = [];

// IMMUTABLE after buildBlocks() — these define the original input structure and
// the final output structure. Merge / unmerge NEVER touches them.
//   paras[i]  = { zhFull, enFull, zhSentences, enSentences }
//   zhSeps[i] / enSeps[i] = literal-\n separator string between paras[i] / paras[i+1]
let paras = [], zhSeps = [], enSeps = [];

// MUTABLE view-layer state. Each group is a visual block on screen.
// Merging adjacent groups concatenates their paraIdxs; unmerging splits back.
//   groups[gi] = { paraIdxs: number[], combined: boolean }
// Initial state: groups.length === paras.length, each holds one paragraph.
let groups = [];
let aiBusy = false;
let aiConfig = {};
let referenceFiles = [];
let enSuggestionRaw = null;
let manualAlignUsed = false;
let lastRawOutput = '';
const clientId = 'client-' + Math.random().toString(36).slice(2) + '-' + Date.now();
let heartbeatTimer = null;

// ── Help toggle ─────────────────────────────────────────────────────────────
function toggleHelp() {
  document.getElementById('helpPanel').classList.toggle('on');
  document.getElementById('helpToggle').classList.toggle('active');
}

// ── Load glossary ───────────────────────────────────────────────────────────
window.addEventListener('load', async () => {
  const tag = document.getElementById('glossTag');
  try {
    const res  = await fetch('/api/glossary');
    const data = await res.json();
    if (data.error) {
      tag.textContent = '⚠ ' + data.error;
      tag.className = 'gloss-tag warn';
    } else {
      glossary = data.terms;
      tag.textContent = '✓ Glossary: ' + glossary.length + ' terms';
      tag.className = 'gloss-tag';
    }
  } catch (e) {
    tag.textContent = '⚠ Failed to load glossary';
    tag.className = 'gloss-tag warn';
  }

  try {
    const res = await fetch('/api/ai/config');
    const data = await res.json();
    if (!data.error) {
      aiConfig = data;
      document.getElementById('aiModelKey').value = data.model_key || 'claude_opus_46';
      document.getElementById('checkPrompt').value = data.check_prompt || '';
      document.getElementById('alignPrompt').value = data.align_prompt || '';
      await loadReferences();
      renderTemplateSelect();
      syncTemplateEditor();
    }
  } catch (e) {
    // Local-only config; ignore load failures and keep defaults in the UI.
  }

  try {
    const res = await fetch('/api/ai/creds');
    const data = await res.json();
    updateCredsTag(data);
  } catch (e) {}
});

function updateCredsTag(data) {
  const tag = document.getElementById('credsTag');
  if (data && data.configured) {
    tag.textContent = '✓ ' + data.account;
    tag.className = 'creds-tag';
  } else {
    tag.textContent = '未配置';
    tag.className = 'creds-tag warn';
  }
}

async function saveCreds() {
  const account = document.getElementById('credsAccount').value.trim();
  const password = document.getElementById('credsPassword').value.trim();
  if (!account || !password) { alert('账号和密码不能为空'); return; }
  try {
    setAiBusy(true, '保存账密中…');
    const res = await fetch('/api/ai/creds', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ account, password })
    });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || '保存失败');
    updateCredsTag({ configured: true, account: data.account });
    document.getElementById('credsPassword').value = '';
    setStatus('账密已保存 ✓', 'ok');
  } catch (e) {
    setStatus(String(e.message || e), 'err');
  } finally {
    setAiBusy(false);
  }
}

async function postJson(url, payload, keepalive = false) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    keepalive
  });
  return res;
}

async function registerClient() {
  try {
    await postJson('/api/client/open', { client_id: clientId });
  } catch (e) {
    // Ignore lifecycle registration failures; the tool can still function.
  }
  heartbeatTimer = setInterval(async () => {
    try {
      await postJson('/api/client/ping', { client_id: clientId });
    } catch (e) {
      // Ignore transient ping failures.
    }
  }, 15000);
}

function closeClientSession() {
  const body = JSON.stringify({ client_id: clientId });
  try {
    if (navigator.sendBeacon) {
      navigator.sendBeacon('/api/client/close', new Blob([body], { type: 'application/json' }));
      return;
    }
  } catch (e) {
    // Fall through to fetch keepalive.
  }
  fetch('/api/client/close', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body,
    keepalive: true
  }).catch(() => {});
}

window.addEventListener('load', registerClient);
window.addEventListener('pagehide', closeClientSession);
window.addEventListener('beforeunload', closeClientSession);

// ── Helpers ─────────────────────────────────────────────────────────────────
function colorize(text) {
  return text.replace(/\[([0-9A-Fa-f]{6})\]([\s\S]*?)\[-\]/g,
    (_, hex, body) => '<span style="color:#' + hex + '">' + body + '</span>');
}

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Convert literal "\n" (backslash + n, two chars) typed by the user inside an
// edit textarea into a real <br> for rendered HTML output. Safe to apply AFTER
// esc()/highlightTerms() because neither emits nor escapes a backslash, so the
// only "\n" sequences left in the html are the user's own.
function nlToBr(html) {
  return html.replace(/\\n/g, '<br>');
}

function autoH(ta) {
  ta.style.height = 'auto';
  ta.style.height = (ta.scrollHeight + 2) + 'px';
}

function setStatus(msg, type) {
  const el = document.getElementById('status');
  el.textContent = msg;
  el.className = 'status ' + type;
}

function setUsageMeta(meta) {
  const box = document.getElementById('usageBox');
  if (!meta) {
    box.className = 'usage-box';
    box.textContent = '';
    return;
  }
  const inputTokens = Number(meta.input_tokens || 0);
  const outputTokens = Number(meta.output_tokens || 0);
  const totalTokens = Number(meta.total_tokens || (inputTokens + outputTokens));
  const cost = meta.estimated_cost_usd;
  const costText = (typeof cost === 'number' && !Number.isNaN(cost))
    ? `$${cost.toFixed(6)}`
    : 'unavailable';
  box.textContent =
    `Model: ${meta.model || '-'} | Input: ${inputTokens} | Output: ${outputTokens} | Total: ${totalTokens} | Estimated cost: ${costText}`;
  box.className = 'usage-box on';
}

function setRawOutput(rawText) {
  lastRawOutput = String(rawText || '');
  const box = document.getElementById('debugBox');
  const ta = document.getElementById('debugRawOutput');
  ta.value = lastRawOutput;
  box.className = lastRawOutput ? 'debug-box on' : 'debug-box';
}

async function copyRawOutput() {
  if (!lastRawOutput) return;
  try {
    await navigator.clipboard.writeText(lastRawOutput);
  } catch (e) {
    const ta = document.getElementById('debugRawOutput');
    ta.focus();
    ta.select();
    document.execCommand('copy');
  }
}

function mergeUsageMeta(firstMeta, secondMeta) {
  if (!firstMeta && !secondMeta) return null;
  if (!firstMeta) return secondMeta;
  if (!secondMeta) return firstMeta;
  const inputTokens = Number(firstMeta.input_tokens || 0) + Number(secondMeta.input_tokens || 0);
  const outputTokens = Number(firstMeta.output_tokens || 0) + Number(secondMeta.output_tokens || 0);
  const totalTokens = Number(firstMeta.total_tokens || 0) + Number(secondMeta.total_tokens || 0);
  const firstCost = typeof firstMeta.estimated_cost_usd === 'number' ? firstMeta.estimated_cost_usd : null;
  const secondCost = typeof secondMeta.estimated_cost_usd === 'number' ? secondMeta.estimated_cost_usd : null;
  return {
    model: secondMeta.model || firstMeta.model || '',
    input_tokens: inputTokens,
    output_tokens: outputTokens,
    total_tokens: totalTokens,
    estimated_cost_usd: firstCost !== null && secondCost !== null ? firstCost + secondCost : null
  };
}

function currentTemplates() {
  return Array.isArray(aiConfig.translate_templates) ? aiConfig.translate_templates : [];
}

function currentTemplateId() {
  return aiConfig.selected_translate_template_id || (currentTemplates()[0] && currentTemplates()[0].id) || '';
}

function currentTemplate() {
  return currentTemplates().find(t => t.id === currentTemplateId()) || currentTemplates()[0] || null;
}

function renderTemplateSelect() {
  const select = document.getElementById('translateTemplateSelect');
  const templateId = currentTemplateId();
  select.innerHTML = currentTemplates().map(t =>
    `<option value="${esc(t.id)}"${t.id === templateId ? ' selected' : ''}>${esc(t.name)}</option>`
  ).join('');
}

function renderReferenceSelect(selectedValue) {
  const select = document.getElementById('referenceSelect');
  const options = ['<option value="">(No reference)</option>'];
  referenceFiles.forEach(name => {
    const selected = name === selectedValue ? ' selected' : '';
    options.push(`<option value="${esc(name)}"${selected}>${esc(name)}</option>`);
  });
  select.innerHTML = options.join('');
}

function syncTemplateEditor() {
  const tpl = currentTemplate();
  if (!tpl) return;
  document.getElementById('templateName').value = tpl.name || '';
  document.getElementById('translatePrompt').value = tpl.prompt || '';
  renderReferenceSelect(tpl.reference_file || '');
}

async function loadReferences() {
  const res = await fetch('/api/references');
  const data = await res.json();
  referenceFiles = Array.isArray(data.files) ? data.files : [];
}

function getAiConfigPayload() {
  const tpl = currentTemplate();
  return {
    model_key: document.getElementById('aiModelKey').value,
    translate_prompt: tpl ? (document.getElementById('translatePrompt').value || tpl.prompt || '') : document.getElementById('translatePrompt').value,
    check_prompt: document.getElementById('checkPrompt').value,
    align_prompt: document.getElementById('alignPrompt').value,
    translate_templates: serializeTemplates(),
    selected_translate_template_id: currentTemplateId()
  };
}

function serializeTemplates() {
  const selectedId = currentTemplateId();
  return currentTemplates().map(t => {
    if (t.id !== selectedId) return t;
    return {
      ...t,
      name: document.getElementById('templateName').value.trim() || t.name,
      prompt: document.getElementById('translatePrompt').value,
      reference_file: document.getElementById('referenceSelect').value
    };
  });
}

function onTemplateChange() {
  aiConfig.translate_templates = serializeTemplates();
  aiConfig.selected_translate_template_id = document.getElementById('translateTemplateSelect').value;
  syncTemplateEditor();
  setUsageMeta(null);
}

function onReferenceChange() {
  const tpl = currentTemplate();
  if (tpl) tpl.reference_file = document.getElementById('referenceSelect').value;
}

function newTemplate() {
  aiConfig.translate_templates = serializeTemplates();
  const id = 'template_' + Date.now();
  const template = {
    id,
    name: 'New Template',
    prompt: document.getElementById('translatePrompt').value || '',
    reference_file: document.getElementById('referenceSelect').value || ''
  };
  aiConfig.translate_templates = [...currentTemplates(), template];
  aiConfig.selected_translate_template_id = id;
  renderTemplateSelect();
  syncTemplateEditor();
}

function deleteTemplate() {
  aiConfig.translate_templates = serializeTemplates();
  const templates = currentTemplates();
  if (templates.length <= 1) {
    alert('At least one translate template must remain.');
    return;
  }
  const targetId = currentTemplateId();
  aiConfig.translate_templates = templates.filter(t => t.id !== targetId);
  aiConfig.selected_translate_template_id = aiConfig.translate_templates[0].id;
  renderTemplateSelect();
  syncTemplateEditor();
}

function setAiBusy(busy, message) {
  aiBusy = busy;
  document.querySelectorAll('button').forEach(btn => {
    if (btn.onclick || btn.getAttribute('onclick')) btn.disabled = busy;
  });
  if (busy && message) setStatus(message, 'info');
}

function findTerms(zhText) {
  return glossary.filter(({zh}) => zhText.includes(zh));
}

async function saveAiConfig({ silent = false } = {}) {
  try {
    if (!silent) setAiBusy(true, 'Saving AI settings...');
    const res = await fetch('/api/ai/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(getAiConfigPayload())
    });
    const data = await res.json();
    if (!res.ok || data.error) throw new Error(data.error || 'Failed to save AI settings');
    aiConfig = data;
    renderTemplateSelect();
    syncTemplateEditor();
    setUsageMeta(null);
    if (!silent) setStatus('AI settings saved ✓', 'ok');
  } catch (e) {
    setStatus(String(e.message || e), 'err');
  } finally {
    if (!silent) setAiBusy(false);
  }
}

async function fetchAiAlignSuggestion(zh, enText) {
  const res = await fetch('/api/ai/align', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      ...getAiConfigPayload(),
      zh_text: zh,
      en_text: enText
    })
  });
  const data = await res.json();
  setRawOutput(data.raw_output || '');
  if (!res.ok || data.error) throw new Error(data.error || 'AI align failed');
  return data;
}

async function runAiAlign() {
  const zh = document.getElementById('zh').value.trim();
  const en = document.getElementById('en').value.trim();
  if (!zh) { alert('Please paste the Chinese source first.'); return; }
  if (!en) { alert('Please paste or generate the English draft first.'); return; }

  try {
    manualAlignUsed = true;
    setAiBusy(true, 'AI aligning...');
    const data = await fetchAiAlignSuggestion(zh, en);
    enSuggestionRaw = data.result || '';
    setUsageMeta(data.usage || null);
    setRawOutput(data.raw_output || '');
    buildBlocks(enSuggestionRaw);
    setStatus('AI align completed ✓', 'ok');
  } catch (e) {
    setStatus(String(e.message || e), 'err');
  } finally {
    setAiBusy(false);
  }
}

async function runAiTranslate() {
  const zh = document.getElementById('zh').value.trim();
  const en = document.getElementById('en').value.trim();
  if (!zh) { alert('Please paste the Chinese source first.'); return; }

  try {
    setAiBusy(true, 'AI translating...');
    const res = await fetch('/api/ai/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ...getAiConfigPayload(),
        zh_text: zh
      })
    });
    const data = await res.json();
    setRawOutput(data.raw_output || '');
    if (!res.ok || data.error) throw new Error(data.error || 'AI translate failed');
    let suggestion = data.result || '';
    let usageMeta = data.usage || null;
    if (!manualAlignUsed && suggestion) {
      setStatus('AI translating... auto aligning suggestion...', 'info');
      const aligned = await fetchAiAlignSuggestion(zh, suggestion);
      suggestion = aligned.result || suggestion;
      usageMeta = mergeUsageMeta(usageMeta, aligned.usage || null);
    }
    enSuggestionRaw = suggestion;
    setUsageMeta(usageMeta);
    buildBlocks(enSuggestionRaw);
    setStatus('AI translation completed ✓', 'ok');
  } catch (e) {
    setStatus(String(e.message || e), 'err');
  } finally {
    setAiBusy(false);
  }
}

async function runAiCheck() {
  const zh = document.getElementById('zh').value.trim();
  const en = document.getElementById('en').value.trim();
  if (!zh) { alert('Please paste the Chinese source first.'); return; }
  if (!en) { alert('Please paste or generate the English draft first.'); return; }

  try {
    setAiBusy(true, 'AI checking...');
    const res = await fetch('/api/ai/check', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        ...getAiConfigPayload(),
        zh_text: zh,
        en_text: en
      })
    });
    const data = await res.json();
    setRawOutput(data.raw_output || '');
    if (!res.ok || data.error) throw new Error(data.error || 'AI check failed');
    let suggestion = data.result || '';
    let usageMeta = data.usage || null;
    if (!manualAlignUsed && suggestion) {
      setStatus('AI checking... auto aligning suggestion...', 'info');
      const aligned = await fetchAiAlignSuggestion(zh, suggestion);
      suggestion = aligned.result || suggestion;
      usageMeta = mergeUsageMeta(usageMeta, aligned.usage || null);
    }
    enSuggestionRaw = suggestion;
    setUsageMeta(usageMeta);
    buildBlocks(enSuggestionRaw);
    setStatus('AI check completed ✓', 'ok');
  } catch (e) {
    setStatus(String(e.message || e), 'err');
  } finally {
    setAiBusy(false);
  }
}

// Wrap matched glossary terms in <span class="term-hit"> while escaping HTML.
// Operates on raw text so we can safely interleave escaped chunks with markup.
// Longer terms win when overlapping (e.g. "首领部队" beats "首领").
function highlightTerms(rawText, terms) {
  if (!terms || !terms.length) return esc(rawText);

  const sorted = [...terms].sort((a, b) => b.zh.length - a.zh.length);
  const matches = [];
  for (const t of sorted) {
    let idx = 0;
    while ((idx = rawText.indexOf(t.zh, idx)) !== -1) {
      const start = idx, end = idx + t.zh.length;
      const overlaps = matches.some(m => start < m.end && end > m.start);
      if (!overlaps) matches.push({ start, end, term: t });
      idx = end;
    }
  }
  matches.sort((a, b) => a.start - b.start);

  let html = '', cursor = 0;
  for (const m of matches) {
    html += esc(rawText.slice(cursor, m.start));
    html += '<span class="term-hit" title="' + esc(m.term.en) + '">'
         +  esc(rawText.slice(m.start, m.end))
         +  '</span>';
    cursor = m.end;
  }
  html += esc(rawText.slice(cursor));
  return html;
}

// Render ZH text with term highlighting + literal \n -> <br> + color codes.
// Pipeline order matters: highlightTerms wraps escaped chunks, nlToBr converts
// any user-typed \n inside the body, and colorize finally wraps [RRGGBB]…[-].
function renderZh(text) {
  return colorize(nlToBr(highlightTerms(text, findTerms(text))));
}

function renderEn(text) {
  return colorize(nlToBr(esc(text || '')));
}

function diffSegments(oldText, newText) {
  const oldStr = String(oldText || '');
  const newStr = String(newText || '');
  if (oldStr === newStr) {
    return {
      changed: false,
      oldPrefix: oldStr,
      oldMid: '',
      oldSuffix: '',
      newPrefix: newStr,
      newMid: '',
      newSuffix: ''
    };
  }

  let prefix = 0;
  const maxPrefix = Math.min(oldStr.length, newStr.length);
  while (prefix < maxPrefix && oldStr[prefix] === newStr[prefix]) prefix++;

  let oldSuffixIdx = oldStr.length - 1;
  let newSuffixIdx = newStr.length - 1;
  while (
    oldSuffixIdx >= prefix &&
    newSuffixIdx >= prefix &&
    oldStr[oldSuffixIdx] === newStr[newSuffixIdx]
  ) {
    oldSuffixIdx--;
    newSuffixIdx--;
  }

  return {
    changed: true,
    oldPrefix: oldStr.slice(0, prefix),
    oldMid: oldStr.slice(prefix, oldSuffixIdx + 1),
    oldSuffix: oldStr.slice(oldSuffixIdx + 1),
    newPrefix: newStr.slice(0, prefix),
    newMid: newStr.slice(prefix, newSuffixIdx + 1),
    newSuffix: newStr.slice(newSuffixIdx + 1)
  };
}

function renderDiffPreview(originalText, suggestedText) {
  const diff = diffSegments(originalText, suggestedText);
  if (!diff.changed) return '';

  const newLine = renderEn(diff.newPrefix)
    + (diff.newMid ? `<span class="diff-new">${renderEn(diff.newMid)}</span>` : '')
    + renderEn(diff.newSuffix);
  const oldLine = renderEn(diff.oldPrefix)
    + (diff.oldMid ? `<span class="diff-old">${renderEn(diff.oldMid)}</span>` : '')
    + renderEn(diff.oldSuffix);

  return `<div class="en-preview-line" data-apply="suggested">${newLine}</div><div class="en-preview-line" data-apply="original">${oldLine}</div>`;
}

// Mask [RRGGBB]...[‐] spans so their internal punctuation doesn't trigger splits.
// Returns masked string (same length, placeholder chars).
function maskColors(text) {
  return text.replace(/\[[0-9A-Fa-f]{6}\][\s\S]*?\[-\]/g,
    m => '\x01'.repeat(m.length));
}

// Split ZH paragraph into sentences, keeping [color]...[‐] spans atomic.
// Terminators: 。！？
function splitZhSentences(text) {
  if (!text.trim()) return [text];
  const masked = maskColors(text);
  const parts = [];
  // \x01 is NOT a terminator; runs of it stay with surrounding text
  const re = /[^。！？]+[。！？]*/g;
  let m;
  while ((m = re.exec(masked)) !== null) {
    const slice = text.slice(m.index, m.index + m[0].length).trim();
    if (slice) parts.push(slice);
  }
  return parts.length ? parts : [text.trim()];
}

// Split EN paragraph into sentences.
// Only split at .!? that is followed by whitespace + an uppercase letter.
// This avoids false splits on decimals (0.1%), abbreviations, color codes, etc.
//
// Two-pass strategy:
//   1) Split on the standard "<.!?> + whitespace + UPPERCASE" boundary.
//   2) Re-glue tiny list-marker fragments (e.g. "1.", "2.", "i.", "I.") onto
//      the next sentence so "1. During the event..." stays as one sentence
//      instead of becoming ["1.", "During the event..."].
function splitEnSentences(text) {
  if (!text.trim()) return [text];
  const masked = maskColors(text);
  const pieces = masked.split(/(?<=[.!?])\s+(?=[A-Z])/);
  const raw = [];
  let cursor = 0;
  for (const piece of pieces) {
    const trimmedMasked = piece.trim();
    if (!trimmedMasked) {
      cursor += piece.length;
      continue;
    }
    const start = masked.indexOf(trimmedMasked, cursor);
    const original = text.slice(start, start + trimmedMasked.length).trim();
    raw.push(original);
    cursor = start + trimmedMasked.length;
  }
  const isListMarker = s => /^[0-9]{1,3}\.$/.test(s) || /^[ivxIVX]{1,4}\.$/.test(s);

  const out = [];
  let pending = '';
  for (const p of raw) {
    if (isListMarker(p)) {
      pending = pending ? pending + ' ' + p : p;
    } else {
      out.push(pending ? pending + ' ' + p : p);
      pending = '';
    }
  }
  if (pending) out.push(pending);
  return out;
}

// Parse raw Crowdin input into non-empty paragraphs PLUS the exact \n separator
// between each adjacent pair. Splitting on literal \n yields tokens; runs of empty
// tokens in between encode multi-\n separators (e.g. "A\n\nB" -> ["A","","B"] = sep "\n\n").
// Returns { paras: string[], seps: string[] } where seps.length === paras.length - 1.
function parseRawWithSeps(raw) {
  const tokens = raw.split('\\n');
  const out = [], seps = [];
  let pendingEmpty = 0;
  for (const tok of tokens) {
    const trimmed = tok.trim();
    if (trimmed) {
      if (out.length > 0) {
        // Separator = (1 boundary \n) + (one extra \n per empty token in between)
        seps.push('\\n'.repeat(1 + pendingEmpty));
      }
      out.push(trimmed);
      pendingEmpty = 0;
    } else {
      pendingEmpty++;
    }
  }
  return { paras: out, seps };
}

// ── Build blocks ─────────────────────────────────────────────────────────────
function buildBlocks(suggestionRaw = null) {
  const zhRaw = document.getElementById('zh').value.trim();
  const enRaw = document.getElementById('en').value.trim();
  if (!zhRaw) { alert('Please paste the Chinese source.'); return; }
  // English draft is OPTIONAL — leave it empty to translate from scratch using
  // the table's per-sentence textareas.

  const zh = parseRawWithSeps(zhRaw);
  const en = parseRawWithSeps(enRaw);
  const suggested = parseRawWithSeps(suggestionRaw !== null ? String(suggestionRaw) : '');
  const n  = Math.max(zh.paras.length, en.paras.length);

  paras = [];
  for (let i = 0; i < n; i++) {
    const zhFull = zh.paras[i] || '';
    const enFull = en.paras[i] || '';
    const enSuggested = suggested.paras[i] || '';
    paras.push({
      zhFull,
      enFull,
      enSuggested,
      zhSentences: splitZhSentences(zhFull),
      enSentences: splitEnSentences(enFull),
      enSuggestedSentences: splitEnSentences(enSuggested)
    });
  }

  // Fill separators for the n-1 gaps. EN can be partial / empty; when no EN
  // separator is available, fall back to the ZH separator so the eventual
  // copied output mirrors the source's \n structure rather than collapsing it.
  zhSeps = [];
  enSeps = [];
  for (let i = 0; i < n - 1; i++) {
    const zSep = zh.seps[i] !== undefined ? zh.seps[i] : '\\n';
    const eSep = en.seps[i] !== undefined ? en.seps[i] : zSep;
    zhSeps.push(zSep);
    enSeps.push(eSep);
  }

  // Reset view-layer: every paragraph starts as its own group
  groups = paras.map((_, i) => ({ paraIdxs: [i], combined: false }));

  document.getElementById('cnt').textContent = n;

  // Show wrapper BEFORE renderBlocks so textareas have non-zero scrollHeight
  // when autoH measures them (otherwise heights collapse to min-height = 1 line
  // and long EN content gets clipped by overflow:hidden).
  const wrap = document.getElementById('blocksWrap');
  wrap.style.display = 'block';
  renderBlocks();
  wrap.scrollIntoView({ behavior: 'smooth', block: 'start' });
  setStatus('Built ' + n + ' paragraph(s) ✓', 'ok');
}

// VIEW-LAYER ONLY. Merge group `gi` with group `gi+1` so they render as one
// visual block. Original paras/seps are untouched -> Copy & Preview output the
// EXACT \n structure of the user's draft regardless of how many merges happen.
function mergeWithNext(gi) {
  if (gi < 0 || gi >= groups.length - 1) return;
  groups[gi].paraIdxs = [...groups[gi].paraIdxs, ...groups[gi + 1].paraIdxs];
  groups[gi].combined = false; // combined-view doesn't make sense for merged groups
  groups.splice(gi + 1, 1);
  renderBlocks();
  const labels = groups[gi].paraIdxs.map(i => i + 1).join(' + ');
  setStatus('View-merged paragraphs ' + labels + ' ✓ (output unchanged)', 'ok');
}

// VIEW-LAYER ONLY. Split a multi-paragraph group back into individual groups.
function unmergeGroup(gi) {
  if (gi < 0 || gi >= groups.length) return;
  const g = groups[gi];
  if (g.paraIdxs.length < 2) return;
  const expanded = g.paraIdxs.map(idx => ({ paraIdxs: [idx], combined: false }));
  groups.splice(gi, 1, ...expanded);
  renderBlocks();
  setStatus('Unmerged ✓', 'ok');
}

function renderBlocks() {
  const container = document.getElementById('blocks');
  container.innerHTML = '';
  groups.forEach((g, gi) => container.appendChild(makeBlock(g, gi)));
  // rAF ensures layout has settled before reading scrollHeight
  requestAnimationFrame(() => {
    container.querySelectorAll('textarea').forEach(autoH);
    container.querySelectorAll('textarea[data-sent]').forEach(syncSentencePreview);
    container.querySelectorAll('textarea[data-combined]').forEach(syncCombinedPreview);
  });
}

// Render one visual block per group. A group with N paragraphs renders all N
// paragraphs' sentences sequentially with a divider between them.
function makeBlock(g, gi) {
  const isMulti = g.paraIdxs.length > 1;
  const isLast  = (gi === groups.length - 1);

  // ── Aggregate glossary hits across all paragraphs in this group (deduped)
  const seen  = new Set();
  const terms = [];
  for (const pIdx of g.paraIdxs) {
    for (const t of findTerms(paras[pIdx].zhFull)) {
      const key = t.zh + '\x01' + t.en;
      if (!seen.has(key)) { seen.add(key); terms.push(t); }
    }
  }
  const termHtml = terms.length
    ? terms.map(t =>
        '<span class="pill">'
      +   '<span class="pill-zh">' + esc(t.zh) + '</span>'
      +   '<span class="pill-eq">=</span>'
      +   esc(t.en)
      + '</span>'
      ).join('')
    : '<span class="no-term">no glossary hits</span>';

  // ── Sentence rows. Walk every paragraph in the group; insert a divider
  //    before each paragraph after the first so users can still see the
  //    original boundary. Each textarea is bound to its ORIGINAL paragraph
  //    index via data-para so edits are attributed correctly even after merge.
  let sentRows = '';
  let rowNum   = 0;
  g.paraIdxs.forEach((pIdx, idxInGroup) => {
    const p = paras[pIdx];
    if (idxInGroup > 0) {
      sentRows += `<div class="para-divider">— Paragraph ${pIdx + 1} —</div>`;
    }
    const rowCount = Math.max(p.zhSentences.length, p.enSentences.length, 1);
    for (let si = 0; si < rowCount; si++) {
      rowNum++;
      const zhVal = p.zhSentences[si] !== undefined ? p.zhSentences[si] : '';
      const enVal = p.enSentences[si] !== undefined ? p.enSentences[si] : '';
      const enSuggestedVal = p.enSuggestedSentences[si] !== undefined ? p.enSuggestedSentences[si] : '';
      sentRows += `<div class="sent-row">
        <div class="sent-num">${rowNum}</div>
        <div class="sent-zh">${zhVal ? renderZh(zhVal) : ''}</div>
        <div class="sent-en"><textarea
          data-para="${pIdx}" data-sent="${si}"
          data-original="${esc(enVal)}"
          data-suggested="${esc(enSuggestedVal)}"
          oninput="autoH(this);syncCombined(${gi});syncSentencePreview(this)">${esc(enVal)}</textarea>
          <div class="en-preview"></div></div>
      </div>`;
    }
  });

  // ── Combined view (works for both single and merged groups). Sentence-level
  //    textareas always exist in the DOM (just hidden behind combined view) and
  //    keep their data-para attribution, so syncSentences distributes re-split
  //    sentences back into the correct ORIGINAL paragraphs by DOM order.
  const zhSentencesAll = g.paraIdxs.flatMap(pIdx => paras[pIdx].zhSentences);
  const zhCombined = zhSentencesAll.length > 1
    ? zhSentencesAll.map((s, si) =>
        (si > 0 ? '<span class="sent-sep"> </span>' : '') + renderZh(s)
      ).join('')
    : renderZh(zhSentencesAll[0] || '');
  const enJoined = g.paraIdxs.map(pIdx => paras[pIdx].enFull).filter(Boolean).join(' ');
  const enSuggestedJoined = g.paraIdxs.map(pIdx => paras[pIdx].enSuggested).filter(Boolean).join(' ');
  const combinedHtml = `
    <div class="combined-view" id="body-comb-${gi}">
      <div class="combined-body">
        <div class="combined-zh">${zhCombined}</div>
        <div class="combined-en">
          <textarea data-combined="1"
            data-original="${esc(enJoined)}"
            data-suggested="${esc(enSuggestedJoined)}"
            oninput="autoH(this);syncSentences(${gi});syncCombinedPreview(this)">${esc(enJoined)}</textarea>
          <div class="en-preview"></div>
        </div>
      </div>
    </div>`;

  // ── Header: label + pills + buttons
  const labelText = isMulti
    ? 'Paragraphs ' + g.paraIdxs.map(i => i + 1).join(' + ')
    : 'Paragraph ' + (g.paraIdxs[0] + 1);
  const unmergeBtn = isMulti
    ? `<button class="btn-merge" onclick="unmergeGroup(${gi})"
         title="Split this merged group back into individual paragraphs">↑ Unmerge</button>`
    : '';
  const mergeBtn = `<button class="btn-merge" onclick="mergeWithNext(${gi})" ${isLast ? 'disabled' : ''}
       title="Merge with next block (view-only — does not change the copied output)">⬇ Merge next</button>`;
  const toggleBtn = `<button class="btn-sm" id="toggle-${gi}" onclick="toggleBlock(${gi})">Merge view ▾</button>`;

  const div = document.createElement('div');
  div.className = 'para-block' + (isMulti ? ' is-merged' : '');
  div.id = 'block-' + gi;
  div.innerHTML = `
    <div class="block-hdr">
      <span class="block-label">${labelText}</span>
      <div class="block-pills">${termHtml}</div>
      ${unmergeBtn}
      ${mergeBtn}
      ${toggleBtn}
    </div>
    <div class="block-body" id="body-sent-${gi}">${sentRows}</div>
    ${combinedHtml}`;

  return div;
}

// ── Toggle a group between sentence-level rows and a single combined textarea.
//    Works for both single-paragraph and merged groups; in merged groups the
//    combined textarea spans all paragraphs and on edit-commit the resulting
//    sentences flow back into the correct ORIGINAL paragraphs by DOM order
//    (each sent-row textarea preserves its data-para attribution).
function toggleBlock(gi) {
  const g = groups[gi];
  if (!g) return;
  g.combined = !g.combined;

  const sentView = document.getElementById('body-sent-' + gi);
  const combView = document.getElementById('body-comb-' + gi);
  const btn      = document.getElementById('toggle-' + gi);

  if (g.combined) {
    const sentTas = sentView.querySelectorAll('textarea');
    const joined  = Array.from(sentTas).map(ta => ta.value).filter(Boolean).join(' ');
    const combTa  = combView.querySelector('textarea');
    combTa.value  = joined;
    syncCombinedPreview(combTa);
    sentView.style.display = 'none';
    combView.style.display = 'block';
    btn.textContent = 'Split view ▴';
    btn.classList.add('active');
    autoH(combTa);
  } else {
    const combTa  = combView.querySelector('textarea');
    const reSplit = splitEnSentences(combTa.value);
    const sentTas = sentView.querySelectorAll('textarea');
    sentTas.forEach((ta, si) => {
      ta.value = reSplit[si] !== undefined ? reSplit[si] : '';
      syncSentencePreview(ta);
    });
    sentView.style.display = 'block';
    combView.style.display = 'none';
    btn.textContent = 'Merge view ▾';
    btn.classList.remove('active');
    sentTas.forEach(autoH);
  }
}

// Sync combined textarea when a sentence textarea changes (single-para groups).
function syncCombined(gi) {
  const combTa = document.querySelector('#body-comb-' + gi + ' textarea');
  if (!combTa) return; // multi-para groups have no combined view
  const sentTas = document.querySelectorAll('#body-sent-' + gi + ' textarea');
  combTa.value = Array.from(sentTas).map(ta => ta.value).filter(Boolean).join(' ');
  autoH(combTa);
  syncCombinedPreview(combTa);
}

// Re-split combined EN into sentence rows when combined textarea changes.
function syncSentences(gi) {
  const combTa  = document.querySelector('#body-comb-' + gi + ' textarea');
  if (!combTa) return;
  const reSplit = splitEnSentences(combTa.value);
  const sentTas = document.querySelectorAll('#body-sent-' + gi + ' textarea');
  sentTas.forEach((ta, si) => {
    ta.value = reSplit[si] !== undefined ? reSplit[si] : '';
    autoH(ta);
    syncSentencePreview(ta);
  });
}

function syncSentencePreview(ta) {
  const preview = ta.parentElement.querySelector('.en-preview');
  if (!preview) return;
  preview.innerHTML = renderDiffPreview(ta.dataset.original || '', ta.dataset.suggested || '');
  preview.style.display = preview.innerHTML ? 'block' : 'none';
  if (preview.innerHTML) {
    preview.querySelectorAll('.en-preview-line').forEach(line => {
      line.onclick = () => applyPreviewChoice(ta, line.dataset.apply);
    });
  }
}

function syncCombinedPreview(ta) {
  const preview = ta.parentElement.querySelector('.en-preview');
  if (!preview) return;
  preview.innerHTML = renderDiffPreview(ta.dataset.original || '', ta.dataset.suggested || '');
  preview.style.display = preview.innerHTML ? 'block' : 'none';
  if (preview.innerHTML) {
    preview.querySelectorAll('.en-preview-line').forEach(line => {
      line.onclick = () => applyPreviewChoice(ta, line.dataset.apply);
    });
  }
}

function applyPreviewChoice(ta, applyType) {
  if (applyType === 'suggested') {
    ta.value = ta.dataset.suggested || '';
  } else {
    ta.value = ta.dataset.original || '';
  }
  autoH(ta);
  if (ta.dataset.combined) {
    const gi = Number((ta.closest('.combined-view') || {}).id?.replace('body-comb-', ''));
    if (!Number.isNaN(gi)) syncSentences(gi);
    syncCombinedPreview(ta);
  } else {
    const block = ta.closest('.para-block');
    const gi = Number((block || {}).id?.replace('block-', ''));
    if (!Number.isNaN(gi)) syncCombined(gi);
    syncSentencePreview(ta);
  }
}

// ── Collect final EN for ORIGINAL paragraph index `pIdx`, regardless of which
//    group currently displays it. Sentence-level textareas are always present
//    in the DOM (even when hidden behind combined view) and are kept in sync
//    by syncSentences, so they are the single source of truth.
function getParaEN(pIdx) {
  const tas = document.querySelectorAll(
    'textarea[data-para="' + pIdx + '"][data-sent]'
  );
  return Array.from(tas).map(ta => ta.value).filter(Boolean).join(' ');
}

// ── Preview modal ────────────────────────────────────────────────────────────
// Always renders the ORIGINAL paragraph structure with the EXACT \n separators
// from the user's draft. Merge / unmerge are pure view-layer operations and
// have NO effect here.
function showPreview() {
  if (!paras.length) { alert('Please build the table first.'); return; }

  const nlCount = sep => (sep.match(/\\n/g) || []).length;

  let zhHtml = '';
  let enHtml = '';
  paras.forEach((p, i) => {
    zhHtml += renderZh(p.zhFull);
    // EN side may contain user-typed color codes [RRGGBB]…[-] AND literal \n.
    // Pipeline: esc -> nlToBr -> colorize so all three get rendered live.
    enHtml += renderEn(getParaEN(i));
    if (i < paras.length - 1) {
      zhHtml += '<br>'.repeat(nlCount(zhSeps[i]));
      enHtml += '<br>'.repeat(nlCount(enSeps[i]));
    }
  });
  document.getElementById('pvZh').innerHTML = zhHtml;
  document.getElementById('pvEn').innerHTML = enHtml;

  document.getElementById('overlay').classList.add('on');
}

function closePreview() {
  document.getElementById('overlay').classList.remove('on');
}

document.getElementById('overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('overlay')) closePreview();
});

// ── Copy: rebuild the result using the ORIGINAL paragraph list and the EXACT
//         \n separators captured from the user's draft. Merge / unmerge are
//         purely visual and never touch this output. Nothing is auto-inserted.
async function doCopy() {
  let result = '';
  paras.forEach((_, i) => {
    result += getParaEN(i);
    if (i < paras.length - 1) result += enSeps[i];
  });

  try {
    await navigator.clipboard.writeText(result);
  } catch {
    const ta = document.createElement('textarea');
    ta.value = result;
    ta.style.cssText = 'position:fixed;opacity:0;top:0;left:0';
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  }
  const toast = document.getElementById('toast');
  toast.classList.add('on');
  setTimeout(() => toast.classList.remove('on'), 3000);
}

function confirmAndCopy() { doCopy(); closePreview(); }
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# HTTP request handler
# ─────────────────────────────────────────────────────────────────────────────
class Handler(http.server.BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass

    def _send_json(self, code: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length) if length > 0 else b"{}"
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _touch_client(self, client_id: str):
        if client_id:
            self.server.touch_client(client_id)

    def _remove_client(self, client_id: str):
        if client_id:
            self.server.remove_client(client_id)

    def do_GET(self):
        if self.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path == "/api/glossary":
            terms, msg = load_glossary()
            if terms:
                self._send_json(200, {"terms": terms, "message": msg})
            else:
                self._send_json(200, {"terms": [], "error": msg})
        elif self.path == "/api/ai/config":
            self._send_json(200, load_ai_config())
        elif self.path == "/api/ai/creds":
            self._send_json(200, load_aichat_creds())
        elif self.path == "/api/references":
            self._send_json(200, {"files": list_reference_files()})
        else:
            self.send_error(404)

    def do_POST(self):
        try:
            data = self._read_json()
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON body"})
            return

        try:
            if self.path == "/api/ai/config":
                config = save_ai_config(data)
                self._send_json(200, {"ok": True, **config})
                return

            if self.path == "/api/ai/creds":
                account = str(data.get("account") or "").strip()
                password = str(data.get("password") or "").strip()
                save_aichat_creds(account, password)
                self._send_json(200, {"ok": True, "account": account})
                return

            if self.path == "/api/client/open":
                client_id = str(data.get("client_id") or "").strip()
                if client_id:
                    self._touch_client(client_id)
                self._send_json(200, {"ok": True})
                return

            if self.path == "/api/client/ping":
                client_id = str(data.get("client_id") or "").strip()
                if client_id:
                    self._touch_client(client_id)
                self._send_json(200, {"ok": True})
                return

            if self.path == "/api/client/close":
                client_id = str(data.get("client_id") or "").strip()
                if client_id:
                    self._remove_client(client_id)
                self._send_json(200, {"ok": True})
                return

            if self.path == "/api/ai/translate":
                zh_text = str(data.get("zh_text") or "").strip()
                if not zh_text:
                    self._send_json(400, {"error": "Chinese source is required"})
                    return

                config = save_ai_config(data)
                terms, _ = load_glossary()
                glossary_prompt = build_glossary_prompt(terms, zh_text)
                templates = config.get("translate_templates") or []
                selected_id = config.get("selected_translate_template_id")
                template = next((t for t in templates if t.get("id") == selected_id), None) or (templates[0] if templates else None)
                if not template:
                    raise ValueError("No translate template found")
                reference_text = load_reference_text(template.get("reference_file"))
                result = call_ai_chat(
                    build_single_prompt_messages(
                        fill_prompt_template(
                            template.get("prompt"),
                            zh_text=zh_text,
                            glossary_prompt=glossary_prompt,
                            reference_text=reference_text,
                        )
                    ),
                    config,
                )
                self._send_json(200, {
                    "result": result["text"],
                    "raw_output": result.get("raw_text", ""),
                    "template_name": template.get("name", ""),
                    "reference_file": template.get("reference_file", ""),
                })
                return

            if self.path == "/api/ai/check":
                zh_text = str(data.get("zh_text") or "").strip()
                en_text = str(data.get("en_text") or "").strip()
                if not zh_text:
                    self._send_json(400, {"error": "Chinese source is required"})
                    return
                if not en_text:
                    self._send_json(400, {"error": "English draft is required"})
                    return

                config = save_ai_config(data)
                terms, _ = load_glossary()
                glossary_prompt = build_glossary_prompt(terms, zh_text)
                result = call_ai_chat(
                    build_single_prompt_messages(
                        fill_prompt_template(
                            config.get("check_prompt"),
                            zh_text=zh_text,
                            en_text=en_text,
                            glossary_prompt=glossary_prompt,
                        )
                    ),
                    config,
                )
                self._send_json(200, {
                    "result": result["text"],
                    "raw_output": result.get("raw_text", ""),
                })
                return

            if self.path == "/api/ai/align":
                zh_text = str(data.get("zh_text") or "").strip()
                en_text = str(data.get("en_text") or "").strip()
                if not zh_text:
                    self._send_json(400, {"error": "Chinese source is required"})
                    return
                if not en_text:
                    self._send_json(400, {"error": "English draft is required"})
                    return

                config = save_ai_config(data)
                terms, _ = load_glossary()
                glossary_prompt = build_glossary_prompt(terms, zh_text)
                zh_parsed = parse_raw_with_seps(zh_text)
                align_messages = build_single_prompt_messages(
                    build_align_json_prompt(
                        zh_text,
                        en_text,
                        glossary_prompt,
                        fill_prompt_template(
                            config.get("align_prompt"),
                            zh_text=zh_text,
                            en_text=en_text,
                            glossary_prompt=glossary_prompt,
                        ),
                    )
                )
                result = call_ai_chat(align_messages, config)
                expected_count = len(zh_parsed["paras"])
                try:
                    aligned_text = parse_align_json_response(
                        result["raw_text"],
                        expected_count,
                        zh_parsed["seps"],
                    )
                except Exception as first_error:
                    raw_first = result.get("raw_text", "")
                    retried = False
                    if _looks_like_truncated_align_json(raw_first):
                        retried = True
                        result_retry = call_ai_chat(align_messages, config, temperature=0.1)
                        try:
                            aligned_text = parse_align_json_response(
                                result_retry["raw_text"],
                                expected_count,
                                zh_parsed["seps"],
                            )
                            result = result_retry
                        except Exception as retry_error:
                            setattr(retry_error, "raw_output", result_retry.get("raw_text", ""))
                            raise ValueError(
                                "AI Align response appears truncated or malformed. Please retry."
                            ) from retry_error
                    if not retried:
                        repair_messages = build_single_prompt_messages(
                            build_align_repair_prompt(
                                zh_text,
                                en_text,
                                glossary_prompt,
                                raw_first,
                                expected_count,
                            )
                        )
                        result_retry = call_ai_chat(repair_messages, config, temperature=0.1)
                        try:
                            aligned_text = parse_align_json_response(
                                result_retry["raw_text"],
                                expected_count,
                                zh_parsed["seps"],
                            )
                            result = result_retry
                        except Exception as retry_error:
                            setattr(retry_error, "raw_output", result_retry.get("raw_text", ""))
                            setattr(first_error, "raw_output", raw_first)
                            raise ValueError(
                                "AI Align could not produce the required paragraph count. Please retry."
                            ) from retry_error
                aligned_parsed = parse_raw_with_seps(aligned_text)
                if len(aligned_parsed["paras"]) == expected_count:
                    fixed_paras = _fix_adjacent_heading_shift(
                        zh_parsed["paras"],
                        aligned_parsed["paras"],
                    )
                    if fixed_paras != aligned_parsed["paras"]:
                        rebuilt = ""
                        for idx, para in enumerate(fixed_paras):
                            rebuilt += para
                            if idx < len(fixed_paras) - 1:
                                rebuilt += zh_parsed["seps"][idx] if idx < len(zh_parsed["seps"]) else "\\n"
                        aligned_text = rebuilt
                self._send_json(200, {
                    "result": aligned_text,
                    "raw_output": result.get("raw_text", ""),
                })
                return

            self._send_json(404, {"error": "Not found"})
        except Exception as e:
            raw_output = ""
            if hasattr(e, "raw_output"):
                raw_output = getattr(e, "raw_output") or ""
            self._send_json(500, {"error": str(e), "raw_output": raw_output})


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
class _Server(http.server.HTTPServer):
    allow_reuse_address = True

    def __init__(self, server_address, handler_class):
        super().__init__(server_address, handler_class)
        self._client_lock = threading.Lock()
        self._clients = {}
        self._shutdown_started = False

    def touch_client(self, client_id: str):
        with self._client_lock:
            self._clients[client_id] = time.time()

    def remove_client(self, client_id: str):
        should_shutdown = False
        with self._client_lock:
            self._clients.pop(client_id, None)
            self._drop_stale_clients_locked()
            should_shutdown = not self._clients and not self._shutdown_started
            if should_shutdown:
                self._shutdown_started = True
        if should_shutdown:
            threading.Thread(target=self._delayed_shutdown, daemon=True).start()

    def _drop_stale_clients_locked(self):
        now = time.time()
        stale_ids = [
            client_id
            for client_id, last_seen in self._clients.items()
            if now - last_seen > CLIENT_STALE_SECONDS
        ]
        for client_id in stale_ids:
            self._clients.pop(client_id, None)

    def _delayed_shutdown(self):
        time.sleep(SHUTDOWN_GRACE_SECONDS)
        with self._client_lock:
            self._drop_stale_clients_locked()
            if self._clients:
                self._shutdown_started = False
                return
        threading.Thread(target=self.shutdown, daemon=True).start()

    def run_client_reaper(self):
        while True:
            time.sleep(10)
            with self._client_lock:
                self._drop_stale_clients_locked()
                should_shutdown = not self._clients and not self._shutdown_started
                if should_shutdown:
                    self._shutdown_started = True
            if should_shutdown:
                self._delayed_shutdown()
                return


if __name__ == "__main__":
    # Kill any process still holding the port
    subprocess.run(f"lsof -ti:{PORT} | xargs kill -9 2>/dev/null; true",
                   shell=True, check=False)
    time.sleep(0.3)

    server = _Server(("127.0.0.1", PORT), Handler)
    print(f"Crowdin Checker  →  http://localhost:{PORT}")
    print(f"Glossary: {os.path.join(SCRIPT_DIR, GLOSSARY_FILE)}")
    print("Press Ctrl+C to stop.\n")

    def _open_browser():
        time.sleep(0.7)
        webbrowser.open(f"http://localhost:{PORT}")

    threading.Thread(target=_open_browser, daemon=True).start()
    threading.Thread(target=server.run_client_reaper, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()
