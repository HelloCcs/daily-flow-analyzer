from __future__ import annotations

import ctypes
import json
import re
from ctypes import wintypes
from urllib.parse import parse_qs, urlparse

from .paths import token_path


class DATA_BLOB(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def normalize_token(value: str) -> str:
    """Accept a raw token, an MCP URL, or the JSON snippet shown by Tushare."""
    text = (value or "").strip()
    if not text:
        return ""
    try:
        parsed = json.loads(text)
        text = json.dumps(parsed, ensure_ascii=False)
    except (TypeError, ValueError):
        pass
    url_match = re.search(r"https?://[^\s\"']+", text)
    if url_match:
        query_token = parse_qs(urlparse(url_match.group(0)).query).get("token", [])
        if query_token:
            text = query_token[0]
    token_match = re.search(r"(?<![0-9a-fA-F])([0-9a-fA-F]{40,})(?![0-9a-fA-F])", text)
    if token_match:
        return token_match.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{32,}", text):
        return text
    raise ValueError("无法识别 Token。请粘贴纯 Token、Tushare MCP URL 或完整 MCP JSON。")


def _blob(data: bytes) -> tuple[DATA_BLOB, ctypes.Array]:
    buffer = ctypes.create_string_buffer(data)
    return DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte))), buffer


def save_token(token: str) -> None:
    token = normalize_token(token)
    if not token:
        token_path().unlink(missing_ok=True)
        return
    source, source_buffer = _blob(token.encode("utf-8"))
    output = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(source), "DailyFlowAnalyzer", None, None, None, 0, ctypes.byref(output)
    ):
        raise ctypes.WinError()
    try:
        encrypted = ctypes.string_at(output.pbData, output.cbData)
        temporary = token_path().with_suffix(".tmp")
        temporary.write_bytes(encrypted)
        temporary.replace(token_path())
    finally:
        ctypes.windll.kernel32.LocalFree(output.pbData)
    del source_buffer


def load_token() -> str:
    path = token_path()
    if not path.exists():
        return ""
    encrypted = path.read_bytes()
    source, source_buffer = _blob(encrypted)
    output = DATA_BLOB()
    if not ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(source), None, None, None, None, 0, ctypes.byref(output)
    ):
        return ""
    try:
        return ctypes.string_at(output.pbData, output.cbData).decode("utf-8")
    finally:
        ctypes.windll.kernel32.LocalFree(output.pbData)
        del source_buffer
