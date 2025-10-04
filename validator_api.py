import os, time, re, json
from flask import Blueprint, request, jsonify
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests

bp = Blueprint("validator_api", __name__, url_prefix="/api")

# 位置规则：前两位数字，第3位字母，第4/5位不得为数字（可为字母或 -），后两位数字，长度=7
CODE_RE = re.compile(r"^\d{2}[A-Za-z][A-Za-z-]{2}\d{2}$")

# 环境变量（Heroku Config Vars）
AUTH_USER = os.environ.get("AUTHORIZATION_USER") or os.environ.get("AUTH_USER")
AUTH_KEY  = os.environ.get("AUTHORIZATION_KEY") or os.environ.get("AUTH_KEY")
CONCURRENCY = int(os.environ.get("VALIDATOR_CONCURRENCY", "2"))  # 默认2
TIMEOUT = float(os.environ.get("FDA_TIMEOUT", "6.0"))
CACHE_TTL = int(os.environ.get("FDA_CACHE_TTL", "43200"))        # 12小时
BASE_URL = "https://www.accessdata.fda.gov/rest/pcbapi/v1/productcode/"

# 简单 TTL 缓存（内存）
_cache = {}  # code -> (value, expire_ts)
def _now(): return time.time()
def _cache_get(code):
    ent = _cache.get(code)
    if not ent: return None
    val, exp = ent
    if exp < _now():
        _cache.pop(code, None)
        return None
    return val
def _cache_set(code, value):
    _cache[code] = (value, _now() + CACHE_TTL)
    if len(_cache) > 10000:
        for i, k in enumerate(list(_cache.keys())):
            if i % 10 == 0:
                _cache.pop(k, None)

def _clean_code(s):  # 统一大写
    return (s or "").strip().upper()

def _hit_fda(code):
    """调用 FDA 校验一个 code，返回 'Valid' / 'Invalid' / ('error','原因')"""
    signature = str(int(time.time()))
    url = f"{BASE_URL}{code}?signature={signature}"
    headers = {
        "Authorization-User": AUTH_USER or "",
        "Authorization-Key": AUTH_KEY or "",
        "User-Agent": "netchb-validator/1.0 (+https://herokuapp.com)"
    }
    try:
        r = requests.get(url, headers=headers, timeout=TIMEOUT)
        data = r.json()
    except Exception:
        return ("error", "request_error")

    rc = data.get("APIRETURNCODE")
    # 403=Valid；404=Invalid；402=长度非法；401/410/411=鉴权异常
    if rc == 403:
        return "Valid"
    elif rc == 404:
        return "Invalid"
    elif rc == 402:
        return "Invalid"  # 长度非法同样视为 Invalid
    elif rc in (401, 410, 411):
        return ("error", "auth_error")
    else:
        return ("error", f"unknown_{rc}")

def _query_with_retries(code):
    """保守重试（最多2次指数退避），并写入缓存"""
    cached = _cache_get(code)
    if cached is not None:
        return code, cached
    delay = 0.6
    for _ in range(3):
        res = _hit_fda(code)
        if isinstance(res, tuple) and res[0] == "error":
            if res[1] in ("request_error",):
                time.sleep(delay); delay *= 2
                continue
            else:
                return code, res
        else:
            _cache_set(code, res)
            return code, res
    return code, ("error", "retry_exceeded")

@bp.route("/validate-codes", methods=["POST"])
def validate_codes():
    if not (AUTH_USER and AUTH_KEY):
        return jsonify({"error": "server_not_configured"}), 500

    payload = request.get_json(silent=True) or {}
    codes = payload.get("codes")
    if not isinstance(codes, list):
        return jsonify({"error": "codes_must_be_list"}), 400

    results, errors = {}, {}

    # 预清洗+本地规则二次校验（双保险）
    cleaned = []
    for c in codes:
        if not isinstance(c, str):
            continue
        x = _clean_code(c)
        if CODE_RE.match(x):
            cleaned.append(x)
        else:
            results[x or c] = "Invalid"

    unique_codes = sorted(set(cleaned))

    with ThreadPoolExecutor(max_workers=CONCURRENCY) as ex:
        futs = [ex.submit(_query_with_retries, code) for code in unique_codes]
        for fu in as_completed(futs):
            code, val = fu.result()
            if isinstance(val, tuple) and val[0] == "error":
                errors[code] = val[1]
            else:
                results[code] = val

    return jsonify({"results": results, "errors": errors})
