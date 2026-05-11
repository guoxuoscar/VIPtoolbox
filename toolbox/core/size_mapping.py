# -*- coding: utf-8 -*-
"""
尺码与字段的外部映射：放在「尺码映射」文件夹里的 JSON，可自行增改，无需改代码。
"""
import json
import os
import re
from typing import Dict, Optional

import sys

from toolbox.core.utils import APP_ROOT, BASE_DIR

# 打包后映射在 exe 旁（由运行时钩子复制）；开发时在 toolbox 下
MAPPINGS_DIR = (
    os.path.join(APP_ROOT, "尺码映射") if getattr(sys, "frozen", False) else os.path.join(BASE_DIR, "尺码映射")
)
SIZE_ALIAS_FILE = os.path.join(MAPPINGS_DIR, "尺码别名.json")
FIELD_EXTRA_FILE = os.path.join(MAPPINGS_DIR, "字段别名补充.json")

# 程序内置的常见尺码写法（外部 JSON 可覆盖同名字段）
_STATIC_SIZE_ALIASES = {
    "2XL": "2XL",
    "2xl": "2XL",
    "XXL": "2XL",
    "xxl": "2XL",
    "1XL": "2XL",
    "1xl": "2XL",
    "3XL": "3XL",
    "3xl": "3XL",
    "XXXL": "3XL",
    "xxxl": "3XL",
    "4XL": "4XL",
    "4xl": "4XL",
    "XXXXL": "4XL",
    "xxxxl": "4XL",
    "5XL": "5XL",
    "5xl": "5XL",
    "XXXXXL": "5XL",
    "xxxxxl": "5XL",
    "XL": "XL",
    "xl": "XL",
    "L": "L",
    "l": "L",
    "M": "M",
    "m": "M",
    "S": "S",
    "s": "S",
    "XS": "XS",
    "xs": "XS",
    "XXS": "XXS",
    "xxs": "XXS",
    "均码": "均码",
    "均": "均码",
}

# 唯品界面勾选用的标准尺码集合
_KNOWN_CANONICAL = frozenset(
    ["XS", "XXS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "均码"]
)


def _norm_token(s: str) -> str:
    return re.sub(r"\s+", "", (s or "").strip()).upper()


def _xxxxxl_to_nxl(t: str) -> Optional[str]:
    """XXL、XXXL 等：连续 X 后以 L 结尾，转为 2XL、3XL…"""
    if not re.fullmatch(r"X+L", t):
        return None
    nx = t.count("X")
    if nx <= 1:
        return "XL"
    return f"{nx}XL"


def _digit_size_code(t: str) -> Optional[str]:
    """两位数字码 07→2XL 等（与 Excel 数字尺码列一致）。"""
    m = {"03": "S", "04": "M", "05": "L", "06": "XL", "07": "2XL", "08": "3XL", "09": "4XL", "10": "5XL"}
    return m.get(t)


def _canonical_one(t: str, external: Optional[Dict[str, str]]) -> Optional[str]:
    if not t:
        return None
    if external and t in external:
        v = external[t]
        return v.upper() if v else None
    if t in _STATIC_SIZE_ALIASES:
        return _STATIC_SIZE_ALIASES[t].upper()
    m = re.fullmatch(r"(\d)XL", t)
    if m:
        return f"{m.group(1)}XL"
    xxl = _xxxxxl_to_nxl(t)
    if xxl:
        return xxl
    dc = _digit_size_code(t)
    if dc:
        return dc
    return None


def canonicalize_size(text: str, external: Optional[Dict[str, str]] = None) -> Optional[str]:
    """
    把图片/OCR/Excel 里各种尺码写法统一成界面用的标准码（S、M、…、2XL、3XL）。
    无法识别时返回 None。
    """
    if text is None:
        return None
    raw = str(text).strip()
    if not raw:
        return None

    parts = re.split(r"[/,，、|]", raw)
    candidates = [raw] + [p for p in parts if p and p.strip()]

    for c in candidates:
        t = _norm_token(c)
        if not t:
            continue
        r = _canonical_one(t, external)
        if r and r in _KNOWN_CANONICAL:
            return r
        if r == "均码":
            return r

        # 云端常见粘连：S9668 / XL11854，优先取开头尺码前缀
        m_prefix = re.match(r"^(XXS|XS|X+L|\dXL|XL|[SML])(?=\d)", t, re.I)
        if m_prefix:
            seg = m_prefix.group(1).upper()
            r3 = _canonical_one(seg, external)
            if r3 and (r3 in _KNOWN_CANONICAL or r3 == "均码"):
                return r3

        m = re.search(
            r"(?:^|[^A-Z0-9\u4e00-\u9fff])(X+L|\dXL|XS|XXS|[SMLX]|均码)(?:$|[^A-Z0-9\u4e00-\u9fff])",
            t,
            re.I,
        )
        if m:
            seg = m.group(1).upper()
            r2 = _canonical_one(seg, external)
            if r2 and r2 in _KNOWN_CANONICAL:
                return r2
            if r2 == "均码":
                return r2

    return None


def _load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def load_external_size_aliases() -> Dict[str, str]:
    """读取 尺码别名.json：键与值都会转成大写无空格后参与匹配。"""
    data = _load_json(SIZE_ALIAS_FILE, {})
    out: Dict[str, str] = {}
    if not isinstance(data, dict):
        return out
    items = data.get("尺码别名", data)
    if not isinstance(items, dict):
        return out
    for k, v in items.items():
        if not isinstance(k, str) or not isinstance(v, str):
            continue
        if k.strip() == "说明":
            continue
        ku = _norm_token(k)
        if ku:
            out[ku] = _norm_token(v) or v.strip()
    return out


def load_external_field_aliases() -> Dict[str, str]:
    """读取 字段别名补充.json：图片上的词 → 表格标准列名。"""
    data = _load_json(FIELD_EXTRA_FILE, {})
    out: Dict[str, str] = {}
    if not isinstance(data, dict):
        return out
    items = data.get("字段别名", data)
    if not isinstance(items, dict):
        return out
    for k, v in items.items():
        if not isinstance(k, str) or not isinstance(v, str):
            continue
        if k.strip() == "说明":
            continue
        out[k.strip()] = v.strip()
    return out


def ensure_default_mapping_files():
    """首次运行时创建文件夹与示例 JSON（已有文件则不动）。"""
    os.makedirs(MAPPINGS_DIR, exist_ok=True)
    if not os.path.exists(SIZE_ALIAS_FILE):
        sample = {
            "说明": "左侧为可能出现的尺码写法，右侧为程序统一后的尺码（须与界面勾选一致）。可自行增删行。",
            "尺码别名": {
                "XXL": "2XL",
                "XXXL": "3XL",
                "XXXXL": "4XL",
                "XXXXXL": "5XL",
            },
        }
        with open(SIZE_ALIAS_FILE, "w", encoding="utf-8") as f:
            json.dump(sample, f, ensure_ascii=False, indent=2)
    if not os.path.exists(FIELD_EXTRA_FILE):
        sample = {
            "说明": "左侧为图片上可能出现的量体名称，右侧为表格里的标准列名。可自行增加。",
            "字段别名": {
                "裤口": "裤脚围",
            },
        }
        with open(FIELD_EXTRA_FILE, "w", encoding="utf-8") as f:
            json.dump(sample, f, ensure_ascii=False, indent=2)
