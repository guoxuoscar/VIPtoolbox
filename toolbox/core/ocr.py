import os, json, subprocess, base64, urllib.request, urllib.parse, tempfile, shutil, re, time

# 默认可用的 PaddleX PP-OCRv5 云端配置（用户要求内置）
DEFAULT_PADDLEX_API_URL = "https://7chf57u39cg2s4n2.aistudio-app.com/ocr"
DEFAULT_PADDLEX_TOKEN = "49652bfd327e8bd1ff7524250f6ea07a6f53255c"

class OCREngine:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        self.ocr_dir = os.path.join(base_dir, "ocr_engine")
        self.paddle_ocr_path = self._find_paddle_ocr()
        self.baidu_config = self._load_baidu_config()

    def _find_paddle_ocr(self):
        search_paths = [
            os.path.join(self.ocr_dir, "PaddleOCR-json.exe"),
            os.path.join(self.ocr_dir, "PaddleOCR_json.exe"),
            os.path.join(self.ocr_dir, "PaddleOCR-json_v1.4.1", "PaddleOCR-json.exe"),
        ]
        for path in search_paths:
            if os.path.exists(path):
                print(f"[OCR] 找到PaddleOCR: {path}")
                return path
        print(f"[OCR] 未找到PaddleOCR")
        return None

    def _load_baidu_config(self):
        config_file = os.path.join(self.ocr_dir, "baidu_api.json")
        if os.path.exists(config_file):
            try:
                with open(config_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        data.setdefault("app_id", "")
                        data.setdefault("api_key", "")
                        data.setdefault("secret_key", "")
                        data.setdefault("paddlex_api_url", "")
                        data.setdefault("paddlex_token", "")
                        data.setdefault("cloud_provider", "auto")
                        if not str(data.get("paddlex_api_url", "")).strip():
                            data["paddlex_api_url"] = DEFAULT_PADDLEX_API_URL
                        if not str(data.get("paddlex_token", "")).strip():
                            data["paddlex_token"] = DEFAULT_PADDLEX_TOKEN
                        if not str(data.get("cloud_provider", "")).strip() or str(data.get("cloud_provider", "")).strip() == "auto":
                            data["cloud_provider"] = "paddlex"
                        return data
            except:
                pass
        return {
            "app_id": "",
            "api_key": "",
            "secret_key": "",
            "paddlex_api_url": DEFAULT_PADDLEX_API_URL,
            "paddlex_token": DEFAULT_PADDLEX_TOKEN,
            "cloud_provider": "paddlex",
        }
    
    def save_baidu_config(self, app_id, api_key, secret_key):
        config_file = os.path.join(self.ocr_dir, "baidu_api.json")
        os.makedirs(self.ocr_dir, exist_ok=True)
        merged = dict(self.baidu_config or {})
        merged.update({"app_id": app_id, "api_key": api_key, "secret_key": secret_key})
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        self.baidu_config = merged

    def save_cloud_config(self, app_id="", api_key="", secret_key="", paddlex_api_url="", paddlex_token="", cloud_provider="auto"):
        config_file = os.path.join(self.ocr_dir, "baidu_api.json")
        os.makedirs(self.ocr_dir, exist_ok=True)
        merged = dict(self.baidu_config or {})
        merged.update(
            {
                "app_id": app_id,
                "api_key": api_key,
                "secret_key": secret_key,
                "paddlex_api_url": paddlex_api_url,
                "paddlex_token": paddlex_token,
                "cloud_provider": cloud_provider or "auto",
            }
        )
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        self.baidu_config = merged

    def is_paddle_ocr_available(self):
        return bool(self.paddle_ocr_path) and os.path.exists(self.paddle_ocr_path)

    def is_baidu_available(self):
        api_key = str(self.baidu_config.get("api_key", "")).strip()
        secret_key = str(self.baidu_config.get("secret_key", "")).strip()
        # 兼容两种模式：
        # 1) 旧版：api_key + secret_key（先换 token）
        # 2) 新版：bce-v3 单 API Key（Authorization: Bearer）
        return bool((api_key and secret_key) or api_key.startswith("bce-v3/"))

    def is_paddlex_available(self):
        cfg = self.baidu_config or {}
        url = str(cfg.get("paddlex_api_url", "")).strip()
        token = str(cfg.get("paddlex_token", "")).strip()
        return bool(url and token)

    def _is_bce_v3_api_key(self):
        api_key = str(self.baidu_config.get("api_key", "")).strip()
        return api_key.startswith("bce-v3/")

    def _prepare_cloud_image(self, image_path):
        """
        云端 OCR 预处理：放大 + 轻度增强，提升小字/浅色图识别率。
        返回 (used_path, temp_path_or_none)。
        """
        temp_path = None
        try:
            from PIL import Image, ImageEnhance
            img = Image.open(image_path).convert("RGB")
            w, h = img.size
            long_side = max(w, h)
            scale = 1.0
            if long_side < 1800:
                scale = 1.6
            elif long_side < 2600:
                scale = 1.25
            if scale > 1.0:
                img = img.resize((int(w * scale), int(h * scale)), Image.Resampling.LANCZOS)
            img = ImageEnhance.Contrast(img).enhance(1.2)
            img = ImageEnhance.Sharpness(img).enhance(1.15)
            temp_path = os.path.join(tempfile.gettempdir(), f"ocr_cloud_{os.getpid()}_{int(time.time()*1000)}.jpg")
            img.save(temp_path, format="JPEG", quality=95, optimize=True)
            return temp_path, temp_path
        except Exception as e:
            print(f"[OCR] 云端预处理失败，回退原图: {e}")
            return image_path, None

    def _post_baidu_ocr(self, endpoint, payload, timeout=30):
        """
        统一请求百度 OCR 接口，兼容 bce-v3 单Key 和 AK/SK token 两种鉴权。
        endpoint 例: general / accurate / general_basic / accurate_basic
        """
        config = self.baidu_config
        url = f"https://aip.baidubce.com/rest/2.0/ocr/v1/{endpoint}"
        data = urllib.parse.urlencode(payload).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        req.add_header("Content-Type", "application/x-www-form-urlencoded")
        if self._is_bce_v3_api_key():
            req.add_header("Authorization", f"Bearer {config.get('api_key', '').strip()}")
        else:
            access_token = self._get_baidu_token(config)
            if not access_token:
                return None
            url = f"{url}?access_token={access_token}"
            req = urllib.request.Request(url, data=data, method="POST")
            req.add_header("Content-Type", "application/x-www-form-urlencoded")
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return json.loads(response.read().decode())

    def _extract_paddlex_items(self, node, items):
        if isinstance(node, dict):
            # PP-OCRv5 常见结构：rec_texts + rec_boxes 成对返回
            rec_texts = node.get("rec_texts")
            rec_boxes = node.get("rec_boxes")
            if isinstance(rec_texts, list):
                for i, txt0 in enumerate(rec_texts):
                    txt = str(txt0 or "").strip()
                    if not txt:
                        continue
                    box = rec_boxes[i] if isinstance(rec_boxes, list) and i < len(rec_boxes) else None
                    x = 0
                    y = 0
                    if isinstance(box, list) and len(box) >= 4 and all(isinstance(v, (int, float)) for v in box[:4]):
                        x = int((box[0] + box[2]) / 2)
                        y = int((box[1] + box[3]) / 2)
                    items.append({"text": txt, "x": x, "y": y})

            txt = node.get("text")
            if txt is None:
                txt = node.get("rec_text")
            if txt is None:
                txt = node.get("transcription")
            if txt is None:
                txt = node.get("block_content")
            if isinstance(txt, str) and txt.strip():
                base_x = 0
                base_y = 0
                box = node.get("bbox") or node.get("box") or node.get("points") or node.get("poly") or node.get("block_bbox")
                if isinstance(box, list) and box:
                    # 兼容 [[x,y],...] 与 [x1,y1,x2,y2] 两类
                    if isinstance(box[0], (list, tuple)) and len(box[0]) >= 2:
                        xs = [int(p[0]) for p in box if isinstance(p, (list, tuple)) and len(p) >= 2]
                        ys = [int(p[1]) for p in box if isinstance(p, (list, tuple)) and len(p) >= 2]
                        if xs and ys:
                            base_x = sum(xs) // len(xs)
                            base_y = sum(ys) // len(ys)
                    elif len(box) >= 4 and all(isinstance(v, (int, float)) for v in box[:4]):
                        base_x = int((box[0] + box[2]) / 2)
                        base_y = int((box[1] + box[3]) / 2)

                # block_content 往往是一整段，拆行拆词生成伪坐标，便于后续表格规则复用
                if "\n" in txt:
                    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
                    for li, line in enumerate(lines):
                        parts = [p for p in re.split(r"\s+", line) if p]
                        if not parts:
                            continue
                        for pi, p in enumerate(parts):
                            # 过滤无意义长串噪声（常见为全 0）
                            if len(p) > 60 and len(set(p)) <= 2:
                                continue
                            items.append({"text": p, "x": base_x + pi * 90, "y": base_y + li * 36})
                else:
                    if len(txt) > 60 and len(set(txt)) <= 2:
                        return
                    items.append({"text": txt.strip(), "x": base_x, "y": base_y})
            for v in node.values():
                self._extract_paddlex_items(v, items)
        elif isinstance(node, list):
            for v in node:
                self._extract_paddlex_items(v, items)

    def ocr_with_paddlex_structured(self, image_path):
        """调用 PaddleX OCR 云端接口，返回统一结构。"""
        if not self.is_paddlex_available():
            return None
        cfg = self.baidu_config or {}
        api_url = str(cfg.get("paddlex_api_url", "")).strip()
        token = str(cfg.get("paddlex_token", "")).strip()
        temp_path = None
        try:
            used_path, temp_path = self._prepare_cloud_image(image_path)
            with open(used_path, "rb") as f:
                img_data = base64.b64encode(f.read()).decode("ascii")
            payload = {
                "file": img_data,
                "fileType": 1,
                "useDocOrientationClassify": False,
                "useDocUnwarping": False,
                "useTextlineOrientation": False,
                # 对 layout-parsing 接口必须打开该开关，否则只做版面不出文字
                "useOcrForImageBlock": True,
            }
            # layout-parsing 默认走版面分析，电商尺码图常出现“识别块为空”，这里强制关闭版面检测只做OCR
            if "/layout-parsing" in api_url:
                payload["useLayoutDetection"] = False
            body = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(api_url, data=body, method="POST")
            req.add_header("Content-Type", "application/json")
            req.add_header("Authorization", f"token {token}")
            with urllib.request.urlopen(req, timeout=45) as response:
                result = json.loads(response.read().decode())
                items = []
                self._extract_paddlex_items(result, items)
                if items:
                    return items
                print(f"[OCR] PaddleX 返回无可解析文字，响应键: {list(result.keys()) if isinstance(result, dict) else type(result)}")
        except Exception as e:
            print(f"[OCR] PaddleX OCR error: {e}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
        return None

    def ocr_with_paddle(self, image_path):
        try:
            paddle_dir = os.path.dirname(self.paddle_ocr_path) if self.paddle_ocr_path else ""
            temp_path = None

            if any(ord(c) > 127 for c in image_path):
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, f"ocr_temp_{os.getpid()}.jpg")
                shutil.copy2(image_path, temp_path)
                image_path = temp_path

            cmd = [self.paddle_ocr_path, "-image_path=" + image_path]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60, encoding="utf-8", cwd=paddle_dir)

            if result.returncode == 0:
                output = result.stdout.strip()
                if output:
                    ansi_escape = re.compile(r"\x1b\[[0-9;]*m")
                    output = ansi_escape.sub("", output)
                    lines = output.split("\n")
                    for line in reversed(lines):
                        line = line.strip()
                        if line.startswith("{") and line.endswith("}"):
                            try:
                                data = json.loads(line)
                                if data.get("code") == 100:
                                    return self._parse_paddle_result(data)
                            except json.JSONDecodeError:
                                continue
            return None
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass

    def _parse_paddle_result(self, data):
        items = []
        for item in data.get("data", []):
            if "text" in item:
                text = item["text"]
                box = item.get("box", [])
                if box and len(box) >= 4:
                    center_x = sum(p[0] for p in box) // 4
                    center_y = sum(p[1] for p in box) // 4
                    items.append({"text": text, "x": center_x, "y": center_y})
                else:
                    items.append({"text": text, "x": 0, "y": 0})
        return items

    def ocr_with_baidu(self, image_path):
        if not self.is_baidu_available():
            return None
        temp_path = None
        try:
            used_path, temp_path = self._prepare_cloud_image(image_path)
            with open(used_path, "rb") as f:
                img_data = base64.b64encode(f.read()).decode()
            payload = {"image": img_data, "language_type": "CHN_ENG", "detect_direction": "true"}
            # 高精度优先，失败再回退通用版
            for endpoint in ("accurate_basic", "general_basic"):
                result = self._post_baidu_ocr(endpoint, payload, timeout=35)
                if not result:
                    continue
                words = result.get("words_result", [])
                if words:
                    texts = [item.get("words") for item in words if item.get("words")]
                    if texts:
                        return texts
                if result.get("error_code"):
                    print(f"[OCR] 百度{endpoint}返回错误: {result.get('error_code')} {result.get('error_msg', '')}")
        except Exception as e:
            print(f"Baidu OCR error: {e}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
        return None

    def ocr_with_baidu_structured(self, image_path):
        """调用百度云端 OCR，并尽量返回带坐标的统一结构。"""
        if not self.is_baidu_available():
            return None
        temp_path = None
        try:
            used_path, temp_path = self._prepare_cloud_image(image_path)
            with open(used_path, "rb") as f:
                img_data = base64.b64encode(f.read()).decode()
            payload = {
                "image": img_data,
                "language_type": "CHN_ENG",
                "detect_direction": "true",
                "probability": "true",
            }
            # 高精度优先，失败再回退通用版
            for endpoint in ("accurate", "general"):
                result = self._post_baidu_ocr(endpoint, payload, timeout=35)
                if not result:
                    continue
                words = result.get("words_result", [])
                items = []
                for item in words:
                    text = str(item.get("words", "")).strip()
                    if not text:
                        continue
                    loc = item.get("location") or {}
                    left = int(loc.get("left", 0) or 0)
                    top = int(loc.get("top", 0) or 0)
                    width = int(loc.get("width", 0) or 0)
                    height = int(loc.get("height", 0) or 0)
                    x = left + width // 2 if width else left
                    y = top + height // 2 if height else top
                    items.append({"text": text, "x": x, "y": y})
                if items:
                    return items
                if result.get("error_code"):
                    print(f"[OCR] 百度{endpoint}返回错误: {result.get('error_code')} {result.get('error_msg', '')}")
        except Exception as e:
            print(f"Baidu structured OCR error: {e}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
        texts = self.ocr_with_baidu(image_path)
        if texts:
            return [{"text": t, "x": 0, "y": 0} for t in texts]
        return None
    
    def _get_baidu_token(self, config):
        try:
            token_url = "https://aip.baidubce.com/oauth/2.0/token"
            params = {"grant_type": "client_credentials", "client_id": config["api_key"], "client_secret": config["secret_key"]}
            url = f"{token_url}?{urllib.parse.urlencode(params)}"
            with urllib.request.urlopen(url, timeout=10) as response:
                result = json.loads(response.read().decode())
                if "access_token" in result:
                    return result["access_token"]
        except Exception as e:
            print(f"获取token异常: {e}")
        return None

    def ocr_image_items(self, image_path, engine="auto"):
        """
        统一 OCR 结果结构，返回 (items, engine_name)。
        items 统一为: [{"text": "...", "x": 123, "y": 456}, ...]
        engine 可选: auto / local / cloud
        """
        mode = str(engine or "auto").strip().lower()
        if mode not in ("auto", "local", "cloud"):
            mode = "auto"

        if mode == "local":
            if self.is_paddle_ocr_available():
                result = self.ocr_with_paddle(image_path)
                if result:
                    return result, "PaddleOCR"
            return None, None

        if mode == "cloud":
            cloud_provider = str((self.baidu_config or {}).get("cloud_provider", "auto")).strip().lower()
            if cloud_provider not in ("auto", "baidu", "paddlex"):
                cloud_provider = "auto"
            if cloud_provider == "paddlex":
                if self.is_paddlex_available():
                    result = self.ocr_with_paddlex_structured(image_path)
                    if result:
                        return result, "PaddleX OCR"
                if self.is_baidu_available():
                    result = self.ocr_with_baidu_structured(image_path)
                    if result:
                        return result, "百度OCR"
            elif cloud_provider == "baidu":
                if self.is_baidu_available():
                    result = self.ocr_with_baidu_structured(image_path)
                    if result:
                        return result, "百度OCR"
                if self.is_paddlex_available():
                    result = self.ocr_with_paddlex_structured(image_path)
                    if result:
                        return result, "PaddleX OCR"
            else:
                if self.is_baidu_available():
                    result = self.ocr_with_baidu_structured(image_path)
                    if result:
                        return result, "百度OCR"
                if self.is_paddlex_available():
                    result = self.ocr_with_paddlex_structured(image_path)
                    if result:
                        return result, "PaddleX OCR"
            return None, None

        # auto: 默认优先本地，失败再云端
        if self.is_paddle_ocr_available():
            result = self.ocr_with_paddle(image_path)
            if result:
                return result, "PaddleOCR"
        if self.is_baidu_available():
            result = self.ocr_with_baidu_structured(image_path)
            if result:
                return result, "百度OCR"
        return None, None

    def ocr_batch_images(self, image_paths, engine="auto", progress_callback=None):
        """
        批量 OCR（用于尺码表测试）。
        返回:
        {
          "total": N,
          "success": M,
          "failed": K,
          "elapsed_sec": 1.23,
          "results": [{"path","success","engine","count","items"}, ...]
        }
        """
        valid_paths = [p for p in (image_paths or []) if p and os.path.isfile(p)]
        start = time.time()
        rows = []
        ok = 0
        fail = 0

        total = len(valid_paths)
        for idx, path in enumerate(valid_paths, start=1):
            items, engine_name = self.ocr_image_items(path, engine=engine)
            is_ok = bool(items)
            if is_ok:
                ok += 1
            else:
                fail += 1
            row = {
                "path": path,
                "success": is_ok,
                "engine": engine_name or "",
                "count": len(items) if items else 0,
                "items": items or [],
            }
            rows.append(row)
            if progress_callback:
                try:
                    progress_callback(idx, total, row)
                except Exception:
                    pass

        return {
            "total": total,
            "success": ok,
            "failed": fail,
            "elapsed_sec": round(time.time() - start, 2),
            "results": rows,
        }
    
    def ocr_image(self, image_path, prefer_local=True):
        if prefer_local:
            return self.ocr_image_items(image_path, engine="auto")
        # 兼容旧逻辑：当 prefer_local=False 时，先云端再本地
        result, source = self.ocr_image_items(image_path, engine="cloud")
        if result:
            return result, source
        return self.ocr_image_items(image_path, engine="local")
