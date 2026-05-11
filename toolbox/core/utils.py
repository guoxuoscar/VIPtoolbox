import os, re, json, random, time, sys
import logging
import winreg


def _compute_paths():
    """
    BASE_DIR：toolbox 包根（打包后在 _MEIPASS/toolbox，内含代码与只读资源）。
    APP_ROOT：可写根目录（exe 同级，配置/日志/u2net/浏览器配置等）。
    """
    here = os.path.abspath(__file__)
    toolbox_pkg = os.path.dirname(os.path.dirname(here))
    if getattr(sys, "frozen", False):
        me = getattr(sys, "_MEIPASS", "") or ""
        bundle = os.path.join(me, "toolbox") if me and os.path.isdir(os.path.join(me, "toolbox")) else me
        if not bundle or not os.path.isdir(bundle):
            bundle = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "_internal", "toolbox")
        app_root = os.path.dirname(os.path.abspath(sys.executable))
        return bundle, app_root
    return toolbox_pkg, os.path.dirname(toolbox_pkg)


BASE_DIR, APP_ROOT = _compute_paths()

# 配置文件：优先 exe/项目根；开发环境若仅有 toolbox 内旧配置则继续用旧路径
CONFIG_FILE = os.path.join(APP_ROOT, "toolbox_config.json")
if not getattr(sys, "frozen", False):
    _legacy = os.path.join(BASE_DIR, "toolbox_config.json")
    if not os.path.isfile(CONFIG_FILE) and os.path.isfile(_legacy):
        CONFIG_FILE = _legacy

REFERENCE_DIR = os.path.join(BASE_DIR, "参考数据")
# 浏览器用户数据必须可写：放在 exe 旁（打包后）；开发时仍在 toolbox 下
PROFILE_DIR = (
    os.path.join(APP_ROOT, ".tb_profile") if getattr(sys, "frozen", False) else os.path.join(BASE_DIR, ".tb_profile")
)
PROXY_FILE = os.path.join(APP_ROOT, "proxies.txt") if getattr(sys, "frozen", False) else os.path.join(BASE_DIR, "proxies.txt")
FIRST_RUN_FLAG = (
    os.path.join(APP_ROOT, ".first_run_done") if getattr(sys, "frozen", False) else os.path.join(BASE_DIR, ".first_run_done")
)

# 尺码表 OCR 档位：ocr_tier.txt 首行 cloud_only=仅云端 | hybrid=本地+云端（抠图始终走本地 u2net，与此无关）
OCR_TIER_FILE = os.path.join(APP_ROOT, "ocr_tier.txt")


def get_ocr_feature_tier():
    """
    返回 hybrid（本地+云端 可切换）或 cloud_only（尺码表只走云端，界面无「自动/仅本地」）。
    与抠图无关，抠图始终为本地 u2net.onnx。
    """
    e = (os.environ.get("TOOLBOX_OCR_TIER") or os.environ.get("OCR_TIER") or "").strip().lower()
    if e in ("cloud_only", "cloud-only", "cloud", "3"):
        return "cloud_only"
    if e in ("hybrid", "all", "local+cloud", "0", "2", "1"):
        return "hybrid"
    try:
        if os.path.isfile(OCR_TIER_FILE):
            with open(OCR_TIER_FILE, "r", encoding="utf-8") as f:
                for raw in f:
                    line = (raw or "").strip()
                    if not line or line.startswith("#"):
                        continue
                    k = line.split()[0].lower() if line else ""
                    if k in ("cloud_only", "cloud-only", "only_cloud", "cloud", "3"):
                        return "cloud_only"
                    if k in ("hybrid", "local+cloud", "0", "1", "2", "all", "default"):
                        return "hybrid"
                    break
    except OSError:
        pass
    return "hybrid"


def get_chrome_path():
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe")
        return winreg.QueryValue(key, "")
    except:
        pass
    return None

CHROME_PATH = get_chrome_path()

# 仅与 Chromium 真实指纹一致：Chrome / Windows x64（避免 Firefox/Safari 与 TLS 不一致）
CHROME_WIN_UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/129.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
]


def get_chrome_ua(randomized=True):
    """返回与 Playwright Chromium 一致的 Chrome Windows UA。"""
    if randomized:
        return random.choice(CHROME_WIN_UAS)
    return CHROME_WIN_UAS[0]

DEFAULT_ANTI_BAN = {
    "enabled": True,
    "min_delay": 3,
    "max_delay": 8,
    "batch_size": 5,
    "batch_rest_min": 30,
    "batch_rest_max": 60,
    "random_ua": True,
    "human_scroll": True,
    "use_proxy": False,
    "auto_retry": True,
    "max_retries": 3,
}

def get_random_ua():
    return get_chrome_ua(True)

def load_anti_ban_config():
    config = load_config()
    saved = config.get("anti_ban", {})
    result = DEFAULT_ANTI_BAN.copy()
    result.update(saved)
    return result

def save_anti_ban_config(anti_ban_config):
    config = load_config()
    config["anti_ban"] = anti_ban_config
    save_config(config)

def load_proxies():
    proxies = []
    if os.path.exists(PROXY_FILE):
        try:
            with open(PROXY_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        if not line.startswith("http"):
                            line = "http://" + line
                        proxies.append(line)
        except:
            pass
    return proxies

def clean(s):
    return re.sub(r'[<>:"/\\|?*\n\r\t]', "_", str(s)).strip()[:40]

CONFIG_VERSION = 2

_CLEANUP_KEYS = frozenset({
    "enhance_dir", "enhance_output", "enhance_mode",
    "enhance_custom_api", "enhance_style_option",
    "enhance_anime_mask", "enhance_inpaint_rect",
    "enhance_api_key", "enhance_secret_key",
    "cutout_engine", "cutout_cloud_api_key", "cutout_cloud_secret_key",
})


def load_config():
    default = {
        "last_excel": "",
        "last_dir": os.path.join(APP_ROOT if getattr(sys, "frozen", False) else BASE_DIR, "下载的商品图片"),
        "resume_enabled": True,
        "concurrency": 1,
        "compress_size": 1200,
        "compress_maxkb": 1024,
        "compress_dir": "",
        "compress_output": "",
        "gen50_maxkb": 1024,
        "gen50_dir": "",
        "gen50_output": "",
        "gen50_scan_depth": 2,
        "compress_naming_mode": "original",
        "compress_naming_custom": "{i}.jpg",
        "gen50_name_as_50": False,
        "cutout_width": 800,
        "cutout_maxkb": 600,
        "cutout_dir": "",
        "cutout_output": "",
        "rename_template": "{n:03d}",
        "rename_dir": "",
        "rename_output": "",
        "operation_history": [],
    }
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                ver = loaded.get("_config_version", 0)
                if ver < 1:
                    for k in _CLEANUP_KEYS:
                        loaded.pop(k, None)
                if ver < 2:
                    loaded.pop("watermark_enable", None)
                loaded["_config_version"] = CONFIG_VERSION
                result = {**default, **loaded}
                return result
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning("加载配置失败 %s: %s", CONFIG_FILE, e)
    return default

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        import logging
        logging.getLogger(__name__).warning("保存配置失败 %s: %s", CONFIG_FILE, e)

def add_operation_history(cfg, operation, details):
    history = cfg.get("operation_history", [])
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    history.insert(0, {"time": timestamp, "operation": operation, "details": details})
    cfg["operation_history"] = history[:50]

import io
from PIL import Image


def _jpeg_best_bytes_under_rgb(img_rgb, max_bytes, min_quality=10, max_quality=95):
    """
    把 RGB 图压成 JPEG：在「不超过 max_bytes」的前提下，取**文件体积最大**的一档（尽量贴近上限）。
    做法：画质从 max_quality 到 min_quality 逐 1 档尝试；并额外试「色度 4:4:4」（subsampling=0），
    同样画质下往往比默认采样更大、更不易出现色块，有利于吃满体积预算。
    说明：若整图本身很简单，即便最高画质体积仍远小于上限，这是 JPEG 极限，只能靠加大像素才能继续涨体积。
    """
    best = None  # (体积, bytes)
    subs_options = (
        ("444", {"subsampling": 0}),
        ("default", {}),
    )
    for _name, extra in subs_options:
        for q in range(max_quality, min_quality - 1, -1):
            buf = io.BytesIO()
            try:
                img_rgb.save(
                    buf,
                    format="JPEG",
                    quality=q,
                    optimize=False,
                    progressive=False,
                    **extra,
                )
            except Exception:
                continue
            sz = len(buf.getvalue())
            if sz <= max_bytes:
                if best is None or sz > best[0]:
                    best = (sz, buf.getvalue())
    if best:
        return best[1]
    # 仍全部超限：打开 optimize 再逐级降（更易压进上限）
    for q in range(max_quality, min_quality - 1, -1):
        buf = io.BytesIO()
        try:
            img_rgb.save(buf, format="JPEG", quality=q, optimize=True, progressive=False)
        except Exception:
            continue
        sz = len(buf.getvalue())
        if sz <= max_bytes:
            return buf.getvalue()
    buf = io.BytesIO()
    img_rgb.save(buf, format="JPEG", quality=min_quality, optimize=True)
    return buf.getvalue()


def compress_to_size(img, filepath, max_kb=1024):
    """压缩 JPEG 到不超过 max_kb；在「不超过」前提下取体积最大的一档（尽量贴近上限）。"""
    if img.mode in ("RGBA", "P", "LA"):
        img = img.convert("RGB")
    elif img.mode != "RGB":
        img = img.convert("RGB")

    target_size = max_kb * 1024
    data = _jpeg_best_bytes_under_rgb(img, target_size)
    with open(filepath, "wb") as f:
        f.write(data)
    return True


def compress_png_to_size(img, max_bytes):
    """压缩PNG图片到指定字节以下，通过调整尺寸"""
    # 转换为RGB（处理透明通道）
    if img.mode == 'RGBA':
        background = Image.new('RGB', img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[3])
        img = background
    
    buf = io.BytesIO()
    img.save(buf, format='PNG')
    current_size = len(buf.getvalue())
    
    if current_size <= max_bytes:
        return img, current_size
    
    # 逐步减小尺寸
    width = img.width
    height = img.height
    step = 0.1  # 每次减少10%
    
    while current_size > max_bytes and width > 100:
        width = int(width * (1 - step))
        height = int(height * (1 - step))
        if width < 100:
            break
        
        resized = img.resize((width, height), Image.LANCZOS)
        buf = io.BytesIO()
        resized.save(buf, format='PNG')
        current_size = len(buf.getvalue())
    
    return resized, current_size


def compress_image_to_size_v2(img, max_kb):
    """
    抠图/透明图保存用：返回 (图片对象, 体积字节)。
    规则：最终按 PNG 计算，体积不超过 max_kb，并尽量接近上限（优先保留更大尺寸）。
    """
    max_bytes = max(1, int(max_kb) * 1024)
    if img.mode != "RGBA":
        img = img.convert("RGBA")

    def _encode_png(_img):
        _buf = io.BytesIO()
        _img.save(_buf, format="PNG", optimize=True, compress_level=9)
        _data = _buf.getvalue()
        return _data, len(_data)

    # 原图已满足上限，直接返回
    data0, size0 = _encode_png(img)
    if size0 <= max_bytes:
        return img, size0

    w, h = img.size
    if w <= 1 or h <= 1:
        return img, size0

    # 二分查找：找「不超过上限」且尽量大的尺寸
    min_side = 32
    low = max(min_side / float(max(w, h)), 0.02)
    high = 1.0
    best_img = None
    best_size = 0

    for _ in range(16):
        mid = (low + high) / 2.0
        nw = max(min_side, int(w * mid))
        nh = max(min_side, int(h * mid))
        cand = img.resize((nw, nh), Image.Resampling.LANCZOS)
        _, csize = _encode_png(cand)
        if csize <= max_bytes:
            if csize > best_size:
                best_img = cand
                best_size = csize
            low = mid
        else:
            high = mid

    # 兜底：极少数图即使很小仍超限，继续强制缩小
    if best_img is None:
        cand = img
        nw, nh = w, h
        while True:
            nw = max(min_side, int(nw * 0.8))
            nh = max(min_side, int(nh * 0.8))
            cand = img.resize((nw, nh), Image.Resampling.LANCZOS)
            _, csize = _encode_png(cand)
            if csize <= max_bytes or (nw <= min_side and nh <= min_side):
                best_img = cand
                best_size = csize
                break

    return best_img, best_size


def find_template_dir() -> str:
    """
    查找表格模板目录，按优先级尝试多个位置：
    1. exe 同级/表格模板/（打包后）
    2. toolbox/表格模板/
    3. ../表格模板/（项目根目录）
    """
    _script_dir = os.path.dirname(os.path.abspath(__file__))
    _toolbox_dir = os.path.dirname(_script_dir)

    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), "表格模板"))
    candidates.append(os.path.join(_toolbox_dir, "表格模板"))
    candidates.append(os.path.join(os.path.dirname(_toolbox_dir), "表格模板"))

    for d in candidates:
        if os.path.isdir(d):
            return d

    raise FileNotFoundError(
        "未找到表格模板目录。尝试了以下路径:\n" +
        "\n".join(f"  - {d}" for d in candidates)
    )


def letterbox_square_white(img, side: int):
    """
    将图等比例缩放后放入 side×side 白底正中（不拉伸），适合转 1:1 主图。
    img 转为 RGB。
    """
    if img.mode not in ("RGB",):
        img = img.convert("RGB")
    w, h = img.size
    if w <= 0 or h <= 0:
        return img
    scale = min(side / w, side / h)
    nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
    small = img.resize((nw, nh), Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (side, side), (255, 255, 255))
    ox = (side - nw) // 2
    oy = (side - nh) // 2
    canvas.paste(small, (ox, oy))
    return canvas
