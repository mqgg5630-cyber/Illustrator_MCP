# -*- coding: utf-8 -*-
"""
agents.common.http_client
统一的 HTTP GET（代理探测 / 重试 / 限速 / 可注入 fetcher 便于离线测试）。
"""

import json
import socket
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request

UA = "AcademicHub/1.0 (mailto:academic_pipeline@research.org)"
DEFAULT_HEADERS = {"User-Agent": UA, "Accept": "application/json, application/xml;q=0.9, */*;q=0.5"}

_ssl_ctx = ssl.create_default_context()


def _detect_proxy():
    for port in (10808, 7890, 7897):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.settimeout(0.2)
            ok = s.connect_ex(("127.0.0.1", port)) == 0
            s.close()
            if ok:
                return f"http://127.0.0.1:{port}"
        except OSError:
            pass
    return None


def _build_opener():
    handlers = [urllib.request.HTTPSHandler(context=_ssl_ctx)]
    proxy = _detect_proxy()
    if proxy:
        handlers.insert(0, urllib.request.ProxyHandler({"http": proxy, "https": proxy}))
    return urllib.request.build_opener(*handlers)


_OPENER = None
_last_call = {}


def _opener():
    global _OPENER
    if _OPENER is None:
        _OPENER = _build_opener()
    return _OPENER


def http_get(url, headers=None, timeout=20, retries=2, min_interval=0.0, rate_key=None):
    """返回 (status_code, bytes)。非 2xx 也返回状态码，不抛异常；网络异常返回 (0, b'')。"""
    if rate_key and min_interval > 0:
        gap = time.time() - _last_call.get(rate_key, 0)
        if gap < min_interval:
            time.sleep(min_interval - gap)
    h = dict(DEFAULT_HEADERS)
    if headers:
        h.update(headers)
    last_status = 0
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers=h)
            with _opener().open(req, timeout=timeout) as resp:
                data = resp.read()
                if rate_key:
                    _last_call[rate_key] = time.time()
                return resp.status, data
        except urllib.error.HTTPError as e:
            last_status = e.code
            if e.code in (404, 400, 401, 403):
                return e.code, b""
            if e.code == 429:
                time.sleep(2 * (attempt + 1))
        except (urllib.error.URLError, socket.timeout, OSError):
            last_status = 0
        time.sleep(0.8 * (attempt + 1))
    return last_status, b""


def get_json(url, **kw):
    status, data = http_get(url, **kw)
    if status != 200 or not data:
        return status, None
    try:
        return status, json.loads(data.decode("utf-8"))
    except ValueError:
        return status, None


def get_text(url, **kw):
    status, data = http_get(url, **kw)
    if status != 200 or not data:
        return status, ""
    return status, data.decode("utf-8", errors="replace")
