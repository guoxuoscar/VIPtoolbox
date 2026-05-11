import os, asyncio, random, threading
import io
from .utils import APP_ROOT, get_random_ua, CHROME_PATH, PROFILE_DIR, compress_image_to_size_v2


def sku_cutout_bytes_to_png(raw, filepath, log=None, min_kb=500, max_kb=600):
    """SKU 图字节 -> 透明 PNG，仅使用本地 u2net.onnx 模型。"""
    log = log or print
    try:
        from PIL import Image
        if len(raw) < 500:
            return False
        model_path = None
        possible_paths = [
            os.path.join(APP_ROOT, "u2net.onnx"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "u2net.onnx"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "toolbox", "u2net.onnx"),
            os.path.join(os.path.expanduser("~"), ".u2net", "u2net.onnx"),
        ]
        for p in possible_paths:
            if os.path.exists(p):
                model_path = p
                break
        if not model_path:
            log("    未找到 u2net.onnx，已跳过抠图")
            return False
        try:
            img = Image.open(io.BytesIO(raw))
            if img.mode != "RGB":
                img = img.convert("RGB")
            from rembg import remove
            log(f"    使用u2net.onnx模型抠图: {model_path}")
            u2net_home = os.environ.get("U2NET_HOME")
            os.environ["U2NET_HOME"] = os.path.dirname(model_path)
            try:
                img_cutout = remove(img)
            finally:
                if u2net_home:
                    os.environ["U2NET_HOME"] = u2net_home
                else:
                    os.environ.pop("U2NET_HOME", None)
        except Exception as e:
            log(f"    rembg抠图失败，使用备用方法: {e}")
            img = Image.open(io.BytesIO(raw)).convert("RGBA")
            pixels = img.load()
            w, h = img.size
            for yy in range(h):
                for x in range(w):
                    r, g, b, a = pixels[x, yy]
                    if r >= 240 and g >= 240 and b >= 240:
                        pixels[x, yy] = (r, g, b, 0)
            img_cutout = img
        # 统一体积策略：不超过目标值，并尽量贴近目标值
        out_img, actual_size = compress_image_to_size_v2(img_cutout, max_kb=max_kb)
        out_img.save(filepath, "PNG", optimize=True, compress_level=9)
        log(f"    抠图完成: {actual_size / 1024:.1f}KB")
        return True
    except Exception as e:
        log(f"  抠图出错: {e}")
        return False


def batch_cutout_skus_under_root(root_dir, log=None):
    """
    对 root_dir 下各款子文件夹内的「SKU图」统一抠透明，写入同款的「透明图」。
    已存在且体积足够的 PNG 会跳过（避免重复算）。
    返回 {"ok","skip","fail"} 计数。
    """
    log = log or print
    ok = skip = fail = 0
    if not root_dir or not os.path.isdir(root_dir):
        return {"ok": 0, "skip": 0, "fail": 0}
    for name in sorted(os.listdir(root_dir)):
        folder = os.path.join(root_dir, name)
        if not os.path.isdir(folder):
            continue
        sku_dir = os.path.join(folder, "SKU图")
        out_dir = os.path.join(folder, "透明图")
        if not os.path.isdir(sku_dir):
            continue
        os.makedirs(out_dir, exist_ok=True)
        for fn in sorted(os.listdir(sku_dir)):
            low = fn.lower()
            if not low.endswith((".jpg", ".jpeg", ".webp")):
                continue
            src = os.path.join(sku_dir, fn)
            base = os.path.splitext(fn)[0]
            dst = os.path.join(out_dir, base + ".png")
            try:
                if os.path.isfile(dst) and os.path.getsize(dst) > 2000:
                    skip += 1
                    continue
                with open(src, "rb") as f:
                    raw = f.read()
                if sku_cutout_bytes_to_png(raw, dst, log=log):
                    ok += 1
                else:
                    fail += 1
            except Exception as e:
                log(f"  统一抠图失败 {name}/{fn}: {e}")
                fail += 1
    log(f"[批量抠图] 完成: 生成 {ok}，跳过 {skip}，失败 {fail}")
    return {"ok": ok, "skip": skip, "fail": fail}


class PW:
    def __init__(self, log=None, captcha_wait=None):
        self.log = log or print
        self._captcha_wait = captcha_wait
        self._loop = self._thread = self._pw = self._ctx = None
        self._ready = threading.Event()
        self._proxy = None
        self._user_agent = None
        self._fingerprint_seed = random.randint(1000, 9999)
        self._headless_mode = True  # 默认无头模式
        # 批量下载时复用同一标签页，避免每款都 new_page 导致窗口反复弹出
        self._reuse_dl_page = None
        self._capture_ref = None
        self._video_list_ref = None

    def set_captcha_wait(self, captcha_wait):
        self._captcha_wait = captcha_wait
    
    def start(self):
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(30)
    
    def _run(self):
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        self._ready.set()
        self._loop.run_forever()
    
    def call(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self._loop).result(300)
    
    def launch(self, headless=True, proxy=None, custom_ua=None):
        # 保存headless设置，用于后续download_product复用
        self._headless_mode = headless
        
        if custom_ua is None:
            final_ua = get_random_ua()
        else:
            final_ua = custom_ua
        final_proxy = proxy
        
        async def _():
            nonlocal final_ua, final_proxy
            from playwright.async_api import async_playwright
            
            max_retries = 3
            last_error = None
            
            for attempt in range(max_retries):
                try:
                    self._pw = await async_playwright().start()
                    os.makedirs(PROFILE_DIR, exist_ok=True)
                    self._user_agent = final_ua
                    self._proxy = final_proxy
                    
                    viewports = [
                        {"width": 1920, "height": 1080},
                        {"width": 1366, "height": 768},
                        {"width": 1440, "height": 900},
                    ]
                    final_viewport = viewports[self._fingerprint_seed % len(viewports)]
                    
                    launch_options = {
                        "headless": headless,
                        "args": [
                            "--no-sandbox",
                            "--disable-dev-shm-usage",
                            "--disable-blink-features=AutomationControlled",
                            "--disable-extensions",
                            "--disable-plugins",
                            "--disable-infobars",
                            "--disable-notifications",
                            "--disable-popup-blocking",
                            "--ignore-certificate-errors",
                        ],
                        "viewport": final_viewport,
                        "user_agent": final_ua,
                        "locale": "zh-CN",
                        "timezone_id": "Asia/Shanghai",
                        "geolocation": {"latitude": 31.2304, "longitude": 121.4737},
                        "permissions": ["geolocation"],
                    }
                    
                    if final_proxy:
                        launch_options["proxy"] = {"server": final_proxy}
                    
                    self._ctx = await self._pw.chromium.launch_persistent_context(
                        PROFILE_DIR, executable_path=CHROME_PATH, **launch_options
                    )
                    self._reuse_dl_page = None
                    self._capture_ref = None
                    self._video_list_ref = None
                    
                    await self._ctx.add_init_script("""
                        Object.defineProperty(navigator, 'webdriver', { get: () => undefined, configurable: true });
                        Object.defineProperty(navigator, 'plugins', { get: () => [{},{},{}], configurable: true });
                        Object.defineProperty(navigator, 'languages', { get: () => ['zh-CN', 'zh', 'en-US', 'en'], configurable: true });
                        Object.defineProperty(navigator, 'platform', { get: () => 'Win32', configurable: true });
                        Object.defineProperty(navigator, 'hardwareConcurrency', { get: () => 8, configurable: true });
                        Object.defineProperty(navigator, 'deviceMemory', { get: () => 8, configurable: true });
                        window.chrome = { runtime: { Oxford: true, lastError: undefined }, loadTimes: function() {}, csi: function() {}, app: {} };
                    """)
                    
                    self.log("[浏览器] 启动成功")
                    break  # 成功则退出重试循环
                    
                except Exception as e:
                    last_error = e
                    self.log(f"[浏览器] 启动失败 (尝试 {attempt+1}/{max_retries}): {str(e)}")
                    
                    # 清理可能的残留进程
                    try:
                        import subprocess
                        subprocess.run(["taskkill", "/F", "/IM", "chrome.exe"], capture_output=True)
                    except:
                        pass
                    
                    # 等待后重试
                    if attempt < max_retries - 1:
                        await asyncio.sleep(2)
                    
                    # 关闭 playwright
                    try:
                        if self._pw:
                            await self._pw.stop()
                    except:
                        pass
            
            if last_error and not self._ctx:
                raise last_error
        
        self.call(_())
    
    def close(self):
        async def _():
            try:
                if self._ctx:
                    await self._ctx.close()
            except:
                pass
            try:
                if self._pw:
                    await self._pw.stop()
            except:
                pass
            self._ctx = self._pw = None
            self._reuse_dl_page = None
            self._capture_ref = None
            self._video_list_ref = None
        
        try:
            self.call(_())
        except:
            pass
    
    def logged_in(self):
        async def _():
            p = await self._ctx.new_page()
            try:
                await p.goto("https://i.taobao.com/", timeout=20000)
                await p.wait_for_timeout(2000)
                return "login" not in p.url
            except:
                return False
            finally:
                await p.close()
        return self.call(_())
    
    def open_login(self, mode="consumer"):
        async def _():
            p = await self._ctx.new_page()
            if mode == "seller":
                # 卖家子账号：先登录，然后强制跳转到淘宝主站避免进入千牛工作台
                await p.goto("https://login.taobao.com/", timeout=60000)
                await p.wait_for_timeout(5000)
                # 等待用户完成登录
                self.log("[登录] 请在浏览器中完成卖家子账号登录...")
                await p.wait_for_timeout(15000)
                # 登录后强制跳转到淘宝主站（避免进入千牛工作台）
                await p.goto("https://www.taobao.com/", timeout=60000)
                await p.wait_for_timeout(3000)
                self.log("[登录] 已跳转到淘宝主站，可以开始下载商品图片")
            elif mode == "qianniu":
                await p.goto("https://work.taobao.com/", timeout=60000)
                await p.wait_for_timeout(3000)
                self.log(f"[登录] 已打开登录页面 (模式：{mode})")
            else:
                await p.goto("https://login.taobao.com/", timeout=60000)
                await p.wait_for_timeout(3000)
                self.log(f"[登录] 已打开登录页面 (模式：{mode})")
        self.call(_())
    
    def download_product(
        self,
        url,
        save_dir,
        code="",
        human_scroll=True,
        matting_mode="none",
        matting_parallel=False,
        matting_workers=2,
        detail_pump_passes=2,
    ):
        """下载商品图片和视频，使用launch时设置的headless模式。
        matting_mode: none=不抠图, each=每款下载时对 SKU 抠透明, batch=仅下载（抠图在全部完成后由界面批量执行）。
        matting_parallel: each 模式下是否并发抠图（边下载边抠，速度更快）。
        matting_workers: each 并发抠图路数（建议 1-4）。
        detail_pump_passes: 仅在详情纵向范围内滑动的遍数（1 或 2），与整页防封滚动无关。"""
        import io
        from PIL import Image
        from urllib.parse import urlparse, unquote
        
        # 保存款号用于视频命名
        product_code = code
        
        async def _():
            saved = {"main": 0, "sku": 0, "detail": 0, "video": 0}
            cutout_tasks = []
            cutout_limit = max(1, min(int(matting_workers or 1), 8))
            cutout_sem = asyncio.Semaphore(cutout_limit)
            os.makedirs(save_dir, exist_ok=True)
            for f in ["主图", "SKU图", "详情图片", "视频", "透明图"]:
                os.makedirs(os.path.join(save_dir, f), exist_ok=True)
            
            captured = {}
            video_list = []
            self._capture_ref = captured
            self._video_list_ref = video_list
            
            reuse = self._reuse_dl_page
            try:
                reuse_ok = reuse is not None and not reuse.is_closed()
            except Exception:
                reuse_ok = False
            if reuse_ok:
                page = reuse
            else:
                page = await self._ctx.new_page()
                self._reuse_dl_page = page
                
                async def on_resp(resp):
                    cap = self._capture_ref
                    vlist = self._video_list_ref
                    if cap is None or vlist is None:
                        return
                    u = resp.url
                    ct = resp.headers.get("content-type", "")
                    is_img = "image" in ct or any(e in u.lower() for e in [".jpg", ".jpeg", ".png", ".webp", ".avif", ".gif"])
                    is_vid = any(e in u.lower() for e in [".mp4", ".m3u8"])
                    if not is_img and not is_vid:
                        return
                    try:
                        body = await resp.body()
                        if is_vid and len(body) > 10000:
                            is_real_video = False
                            if u.lower().endswith(".mp4"):
                                if body[:4] in [b"ftyp", b"mdat", b"moov"] or b"ftyp" in body[:100]:
                                    is_real_video = True
                            if is_real_video:
                                vlist.append((u, body))
                        elif is_img and len(body) > 500:
                            base_url = u.split("?")[0]
                            cap[u] = body
                            cap[base_url] = body
                    except Exception:
                        pass
                
                page.on("response", on_resp)
            
            async def human_scroll_smooth(pg, total_height):
                current_y = 0
                while current_y < total_height:
                    step = random.randint(300, 600)
                    current_y = min(current_y + step, total_height)
                    await pg.evaluate(f"window.scrollTo({{top: {current_y}, behavior: 'smooth'}})")
                    await pg.wait_for_timeout(random.uniform(100, 400))

            async def human_scroll_fast(pg):
                th = await pg.evaluate("document.body.scrollHeight")
                await pg.evaluate(f"window.scrollTo({{top: {th}, behavior: 'auto'}})")
                await pg.wait_for_timeout(800)
            
            async def pump_detail_vertical_band(pg, ds, de, passes=2, fast_mode=False):
                """仅在详情纵向范围内滑动 passes 次（默认2次），兼顾懒加载与速度；不整页慢扫。"""
                try:
                    ds = float(ds)
                    de = float(de)
                except (TypeError, ValueError):
                    return
                if de <= ds + 200:
                    return
                top = max(0, int(ds - 400))
                bot = int(de + 500)
                np = max(1, int(passes))
                step_y = 220 if fast_mode else 185
                wait_scroll = (38 + random.randint(0, 22)) if fast_mode else (68 + random.randint(0, 45))
                wait_band = (95 + random.randint(0, 50)) if fast_mode else (160 + random.randint(0, 90))
                max_ifr = 6 if fast_mode else 12
                for pi in range(np):
                    y = top
                    while y <= bot:
                        await pg.evaluate(f"window.scrollTo(0, {y})")
                        await pg.wait_for_timeout(wait_scroll)
                        y += step_y
                    sh_band = int(await pg.evaluate("document.body.scrollHeight") or 0)
                    await pg.evaluate(f"window.scrollTo(0, {min(bot, sh_band)})")
                    await pg.wait_for_timeout(wait_band)
                    nifr = int(await pg.evaluate("document.querySelectorAll('iframe').length") or 0)
                    for idx in range(min(nifr, max_ifr)):
                        try:
                            await pg.evaluate(
                                """(ix) => {
                                    const fr = document.querySelectorAll('iframe')[ix];
                                    if (!fr) return;
                                    try {
                                        const d = fr.contentDocument;
                                        const w = fr.contentWindow;
                                        if (!d || !d.body || !w) return;
                                        const h = Math.max(d.body.scrollHeight, d.documentElement.scrollHeight, 200);
                                        for (let yy = 0; yy <= h; yy += 150) w.scrollTo(0, yy);
                                        w.scrollTo(0, h);
                                    } catch (e) {}
                                }""",
                                idx,
                            )
                            await pg.wait_for_timeout((32 + random.randint(0, 18)) if fast_mode else (52 + random.randint(0, 35)))
                        except Exception:
                            pass
                    if pi + 1 < np:
                        await pg.wait_for_timeout((48 + random.randint(0, 28)) if fast_mode else (85 + random.randint(0, 45)))
            
            async def check_captcha():
                captcha_indicators = ["访问过于频繁", "系统繁忙", "操作频繁", "亲，访问太频繁了", "账号存在风险", "人机验证"]
                try:
                    if any(ind in page.url.lower() for ind in ["captcha", "verify", "chk_user"]):
                        return True
                    page_content = await page.content()
                    match_count = sum(1 for ind in captcha_indicators if ind in page_content)
                    if match_count >= 2:
                        return True
                except:
                    pass
                return False

            async def resolve_captcha_if_needed(target_url):
                if not await check_captcha():
                    return
                if self._captcha_wait:
                    self.log("[防封] 检测到验证/风控页面，请在浏览器中完成验证，完成后在弹窗中点击确定继续…")
                    loop = asyncio.get_event_loop()
                    await loop.run_in_executor(None, self._captcha_wait)
                    await page.goto(target_url, timeout=60000, wait_until="domcontentloaded")
                    await page.wait_for_timeout(2000)
                    if await check_captcha():
                        if page is not self._reuse_dl_page:
                            await page.close()
                        raise Exception("CAPTCHA_DETECTED")
                else:
                    if page is not self._reuse_dl_page:
                        await page.close()
                    raise Exception("CAPTCHA_DETECTED")
            
            try:
                await page.wait_for_timeout(random.uniform(1, 3))
                
                # 30%概率先访问首页
                if random.random() < 0.3:
                    try:
                        await page.goto("https://www.taobao.com/", timeout=15000, wait_until="domcontentloaded")
                        await page.wait_for_timeout(random.uniform(1, 2))
                    except:
                        pass
                
                target_url = url
                await page.goto(target_url, timeout=60000, wait_until="domcontentloaded")
                
                if "login" in page.url:
                    if page is self._reuse_dl_page:
                        try:
                            await page.goto("about:blank")
                        except Exception:
                            pass
                    else:
                        await page.close()
                    return None
                
                await resolve_captcha_if_needed(target_url)
                
                await page.wait_for_timeout(2000)
                
                await resolve_captcha_if_needed(target_url)
                
                total_h = await page.evaluate("document.body.scrollHeight")
                if human_scroll:
                    await human_scroll_smooth(page, total_h)
                else:
                    await human_scroll_fast(page)
                
                # 点击SKU选项
                try:
                    sku_items_click = await page.query_selector_all('[class*="skuValueWrap"] [class*="valueItem"]')
                    for it in sku_items_click[:12]:
                        try:
                            await it.click()
                            await page.wait_for_timeout(random.uniform(200, 500))
                        except:
                            pass
                except:
                    pass
                
                await page.evaluate("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'})")
                await page.wait_for_timeout(1300)
                
                await resolve_captcha_if_needed(target_url)
                
                # 通过JavaScript获取视频URL
                try:
                    video_info = await page.evaluate("""
                        () => {
                            const urls = [];
                            document.querySelectorAll('video').forEach(v => {
                                if (v.src && v.src.includes('video')) urls.push(v.src);
                            });
                            document.querySelectorAll('[class*="video"], [id*="video"]').forEach(el => {
                                ['data-video', 'data-src', 'data-url', 'data-videourl'].forEach(attr => {
                                    const val = el.getAttribute(attr);
                                    if (val && val.includes('http') && (val.includes('.mp4') || val.includes('.m3u8'))) {
                                        urls.push(val);
                                    }
                                });
                            });
                            document.querySelectorAll('script').forEach(s => {
                                const text = s.textContent || '';
                                const matches = text.match(/https?[^"'<>]+\\.(?:mp4|m3u8)/gi);
                                if (matches) matches.forEach(m => urls.push(m));
                            });
                            return urls;
                        }
                    """)
                    for vid_url in video_info or []:
                        if vid_url and len(vid_url) > 20:
                            video_list.append((vid_url, b"VIDEO_PLACEHOLDER"))
                except:
                    pass
                
                # 鼠标悬停触发视频加载
                self.log("  尝试触发视频加载...")
                video_found = False
                try:
                    main_img_selectors = [
                        '[class*="mainPic"]',
                        '[class*="tb-gallery"]',
                        '[class*="gallery"]',
                        "#J_UlThumb li img",
                        '[class*="video"]',
                    ]
                    for sel in main_img_selectors:
                        try:
                            main_img = await page.query_selector(sel)
                            if main_img:
                                self.log(f"  悬停主图: {sel}")
                                await main_img.hover()
                                await page.wait_for_timeout(950)
                                video_info2 = await page.evaluate("""
                                    () => {
                                        const videos = [];
                                        document.querySelectorAll('video').forEach(v => {
                                            if (v.src && (v.src.includes('video') || v.src.includes('.mp4') || v.src.includes('.m3u8'))) {
                                                videos.push(v.src);
                                            }
                                        });
                                        return videos;
                                    }
                                """)
                                for video_url in video_info2:
                                    if video_url and len(video_url) > 20:
                                        self.log(f"  发现视频: {video_url[:60]}...")
                                        video_list.append((video_url, b"VIDEO_PLACEHOLDER"))
                                        video_found = True
                                if video_found:
                                    break
                        except:
                            continue
                except Exception as e:
                    self.log(f"  视频触发失败: {e}")
                
                await page.evaluate("window.scrollTo({top: 0, behavior: 'auto'})")
                await page.wait_for_timeout(350 + random.randint(0, 150))
                
                await resolve_captcha_if_needed(target_url)
                
                self.log(f"  拦截到{len(captured)}个图片 {len(video_list)}个视频")
                
                def save_jpg(raw, filepath, compress=True, max_kb=1024):
                    try:
                        if len(raw) < 500:
                            return False
                        img = Image.open(io.BytesIO(raw))
                        if compress:
                            from .utils import compress_to_size
                            return compress_to_size(img, filepath, max_kb)
                        if img.mode in ("RGBA", "P", "LA"):
                            img = img.convert("RGB")
                        elif img.mode != "RGB":
                            img = img.convert("RGB")
                        img.save(filepath, "JPEG", quality=95)
                        return True
                    except:
                        return False
                
                async def collect_dom():
                    return await page.evaluate("""() => {
                    const mainImgs=[], skuItems=[], allImgs=[], iframeImgs=[];
                    const seen = new Set();
                    const seenIf = new Set();
                    const cdnDomains = ['alicdn.com', 'taobao.com', 'tmall.com', 'mmall.com', 'etao.com',
                        'gw.alicdn.com', 'img.alicdn.com', 'sc01.alicdn.com', 'sc02.alicdn.com', 'alicdn.net'];
                    const isValidCDN = (url) => url && cdnDomains.some(d => url.includes(d));
                    const getImgUrl = (img) => {
                        let u = img.getAttribute('data-ks-lazyload') || img.getAttribute('data-lazy-src')
                            || img.getAttribute('data-original') || img.getAttribute('data-src')
                            || img.getAttribute('data-img') || img.getAttribute('data-lazyload') || '';
                        if (!u || u.startsWith('data:')) u = img.getAttribute('src') || img.currentSrc || '';
                        if (!u || u.startsWith('data:')) return '';
                        const ss = img.getAttribute('srcset');
                        if (ss) {
                            let best = '', bestW = -1;
                            ss.split(',').forEach(part => {
                                const bits = part.trim().split(/\\s+/);
                                const cand = bits[0];
                                if (!cand || cand.startsWith('data:')) return;
                                let w = 0;
                                bits.forEach(b => { const m = b.match(/^(\\d+)w$/); if (m) w = parseInt(m[1], 10); });
                                if (w > bestW) { bestW = w; best = cand; }
                            });
                            if (best) u = best;
                        }
                        u = (u || '').trim();
                        if (u.startsWith('//')) u = 'https:' + u;
                        return u;
                    };
                    const pushIframe = (u, nw, nh) => {
                        if (!u || u.startsWith('data:') || u.length < 12) return;
                        if (!isValidCDN(u)) return;
                        if (seenIf.has(u)) return;
                        seenIf.add(u);
                        iframeImgs.push({url: u, nw: nw || 0, nh: nh || 0});
                    };
                    
                    document.querySelectorAll('[class*="thumbnailItem"] img').forEach(img => {
                        let u = getImgUrl(img);
                        if(u && !seen.has(u) && isValidCDN(u)) { seen.add(u); mainImgs.push(u); }
                    });
                    document.querySelectorAll('[class*="skuValueWrap"] [class*="valueItem"]').forEach(item => {
                        const img = item.querySelector('img');
                        if(!img) return;
                        let u = getImgUrl(img);
                        if(!u || seen.has(u)) return;
                        seen.add(u);
                        let name = '';
                        const t = item.querySelector('[class*="valueTitle"],[class*="name"],span');
                        if(t) name = t.textContent.trim();
                        if(!name) name = item.getAttribute('title')||item.textContent.trim().substring(0,30);
                        skuItems.push({url:u, name:name.replace(/\\s+/g,' ').trim()});
                    });
                    
                    const isReviewTree = (el) => {
                        let n = el;
                        for (let i = 0; i < 14 && n; i++) {
                            const id = (n.id || '').toLowerCase();
                            const cls = (n.className && String(n.className).toLowerCase()) || '';
                            if (id.includes('review') || id.includes('rate-') || id.includes('comment') ||
                                cls.includes('reviews') || cls.includes('rate-grid') || cls.includes('rategrid') ||
                                cls.includes('comment-') || cls.includes('askitem') || cls.includes('ask-list') ||
                                cls.includes('买家秀') || cls.includes('问大家')) return true;
                            n = n.parentElement;
                        }
                        return false;
                    };
                    document.querySelectorAll('img').forEach(img => {
                        let u = getImgUrl(img);
                        if(!u || u.startsWith('data:') || u.length < 12) return;
                        if (isReviewTree(img)) return;
                        const rect = img.getBoundingClientRect();
                        const w = Math.round(rect.width);
                        const h = Math.round(rect.height);
                        const y = Math.round(rect.top + window.scrollY);
                        const nw = img.naturalWidth;
                        const nh = img.naturalHeight;
                        allImgs.push({url:u, w, h, y, nw, nh});
                    });
                    
                    document.querySelectorAll('iframe').forEach(fr => {
                        try {
                            const d = fr.contentDocument;
                            if(!d) return;
                            d.querySelectorAll('img').forEach(img => {
                                let u = getImgUrl(img);
                                if(!u || u.startsWith('data:')) return;
                                if (isReviewTree(img)) return;
                                const nw = img.naturalWidth, nh = img.naturalHeight;
                                if(nw < 50 && nh < 50) return;
                                pushIframe(u, nw, nh);
                            });
                        } catch(e) {}
                    });
                    
                    let detailStart = 0, detailEnd = 0;
                    let foundStart = false, foundEnd = false;
                    const pageHeight = document.body.scrollHeight;
                    
                    const detailSelectors = [
                        '#J_DetailSection',
                        '[id*="description"]',
                        '[class*="description"]',
                        '[class*="productDetail"]',
                        '[class*="detail"]'
                    ];
                    
                    for(const sel of detailSelectors) {
                        const el = document.querySelector(sel);
                        if(el) {
                            const rect = el.getBoundingClientRect();
                            if(rect.height > 200) {
                                detailStart = rect.top + window.scrollY;
                                foundStart = true;
                                break;
                            }
                        }
                    }
                    
                    if(!foundStart) {
                        const detailTexts = ['图文详情', '详情介绍', '商品详情', '产品详情', '详细描述', 'description'];
                        document.querySelectorAll('*').forEach(el => {
                            if(foundStart) return;
                            const t = el.textContent.trim();
                            const rect = el.getBoundingClientRect();
                            const sy = rect.top + window.scrollY;
                            if(sy < 500) return;
                            if(t.length < 2 || t.length > 20) return;
                            for(const dt of detailTexts) {
                                if(t.includes(dt) || t === dt) {
                                    if(el.tagName.match(/^(DIV|P|SPAN|H[1-6]|A)$/)) {
                                        detailStart = sy + 30;
                                        foundStart = true;
                                        break;
                                    }
                                }
                            }
                        });
                    }
                    
                    if(!foundStart) {
                        const priceSelectors = ['[class*="price"]', '[class*="originalPrice"]', '.tb-detail-price'];
                        for(const sel of priceSelectors) {
                            const el = document.querySelector((sel));
                            if(el) {
                                const rect = el.getBoundingClientRect();
                                detailStart = rect.bottom + window.scrollY;
                                foundStart = true;
                                break;
                            }
                        }
                    }
                    
                    const endTexts = ['本店推荐', '看了又看', '相似推荐', '相关推荐', '猜你喜欢', '精品推荐', '热门推荐', '同类推荐', '其他推荐', '店铺推荐', '为你推荐', '更多推荐',
                        '宝贝评价', '问大家', '全部评价', '用户评价', '累计评价', '评论(', '买家秀', '大家评', '口碑'];
                    document.querySelectorAll('*').forEach(el => {
                        if(!foundStart || foundEnd) return;
                        const t = el.textContent.trim();
                        const rect = el.getBoundingClientRect();
                        const sy = rect.top + window.scrollY;
                        if(sy < detailStart + 300) return;
                        for(const et of endTexts) {
                            if(t.includes(et)) {
                                detailEnd = sy - 30;
                                foundEnd = true;
                                return;
                            }
                        }
                    });
                    
                    const shFull = Math.max(
                        pageHeight,
                        document.body.scrollHeight || 0,
                        document.documentElement.scrollHeight || 0
                    );
                    if(!foundEnd && foundStart) {
                        // 旧逻辑在 detailStart 被误匹配到页面底部时，会与 0.9*pageHeight 重合，只剩约 500px 高「假详情带」
                        detailEnd = Math.max(detailStart + 2000, shFull - 200);
                    }
                    
                    if(!foundStart) {
                        detailStart = pageHeight * 0.15;
                        detailEnd = pageHeight * 0.90;
                    }
                    
                    // 详情起点落在页面下半部过半时多为误匹配（如先命中底部含 detail 字样的模块）
                    if (foundStart && detailStart > shFull * 0.52) {
                        detailStart = shFull * 0.16;
                        detailEnd = Math.max(detailEnd, shFull * 0.92);
                    }
                    // 详情纵向范围过窄时强制延伸到底部附近
                    if (foundStart && (detailEnd - detailStart) < 1000) {
                        detailEnd = Math.max(detailEnd, shFull - 200);
                    }
                    
                    return {mainImgs, skuItems, allImgs, detailStart, detailEnd, iframeImgs};
                }""")
                
                dom_info = await collect_dom()
                _ds0 = int(dom_info.get("detailStart") or 0)
                _de0 = int(dom_info.get("detailEnd") or 0)
                _passes = max(1, min(2, int(detail_pump_passes) if detail_pump_passes else 2))
                if _de0 > _ds0 + 500:
                    await pump_detail_vertical_band(
                        page, _ds0, _de0, passes=_passes, fast_mode=(_passes <= 1)
                    )
                    await page.evaluate("window.scrollTo(0,0)")
                    await page.wait_for_timeout(280 + random.randint(0, 120))
                    dom_info = await collect_dom()
                    self.log(f"  详情区二次采集（懒加载，详情滑{_passes}遍）")
                
                def canon_img_url(u):
                    if not u or not isinstance(u, str):
                        return u
                    u = u.strip()
                    if u.startswith("//"):
                        return "https:" + u
                    return u
                
                def _url_match_key(u):
                    if not u:
                        return ""
                    try:
                        p = urlparse(u)
                        return f"{(p.netloc or '').lower()}{unquote(p.path or '')}"
                    except Exception:
                        return (u or "").split("?")[0]
                
                def detail_size_ok(img, img_url):
                    """详情区：已解码大图用 natural 尺寸；懒加载未解码时允许大占位框（后续用 URL 拉取）。"""
                    nw = int(img.get("nw", 0) or 0)
                    nh = int(img.get("nh", 0) or 0)
                    w = int(img.get("w", 0) or 0)
                    h = int(img.get("h", 0) or 0)
                    if w < 48 or h < 48:
                        return False
                    if nw > 12 and nh > 12:
                        ratio = nw / float(nh) if nh else 1
                        if ratio < 0.08 or ratio > 14:
                            return False
                        if nw >= 100 and nh >= 100:
                            return True
                    if w >= 180 and h >= 180:
                        return True
                    return False
                
                main_urls = [canon_img_url(u) for u in (dom_info["mainImgs"] or []) if u]
                sku_items = dom_info["skuItems"] or []
                for _it in sku_items:
                    if isinstance(_it, dict) and _it.get("url"):
                        _it["url"] = canon_img_url(_it["url"])
                all_imgs = dom_info["allImgs"] or []
                for _im in all_imgs:
                    if isinstance(_im, dict) and _im.get("url"):
                        _im["url"] = canon_img_url(_im["url"])
                ds = dom_info.get("detailStart", 0)
                de = dom_info.get("detailEnd", 0)
                self.log(f"  详情区: {ds} ~ {de}")
                
                valid_cdn_patterns = [
                    "alicdn.com", "alicdn.net", "taobao.com", "tmall.com", "mmall.com",
                    "gw.alicdn.com", "img.alicdn.com", "sc01.alicdn.com", "sc02.alicdn.com",
                ]
                exclude_patterns = ["s.gif", "1x1", "pixel", "tracking", "icon", "logo", "qrcode", "qr-code", "barcode"]
                
                detail_imgs = []
                for img in all_imgs:
                    y = img["y"]
                    img_url = img["url"]
                    
                    if y <= ds or y > de:
                        continue
                    if not detail_size_ok(img, img_url):
                        continue
                    if any(p in img_url.lower() for p in exclude_patterns):
                        continue
                    if not any(d in img_url for d in valid_cdn_patterns):
                        continue
                    detail_imgs.append(img)
                
                detail_imgs.sort(key=lambda x: x["y"])
                detail_urls = []
                seen_urls = set()
                for u in [img["url"] for img in detail_imgs]:
                    if u not in seen_urls:
                        seen_urls.add(u)
                        detail_urls.append(u)
                
                main_set = set(main_urls)
                sku_set = {it["url"] for it in sku_items if isinstance(it, dict) and it.get("url")}
                for entry in dom_info.get("iframeImgs") or []:
                    if isinstance(entry, dict):
                        iu = canon_img_url(entry.get("url", ""))
                    else:
                        iu = canon_img_url(str(entry))
                    if not iu or iu in seen_urls or iu in main_set or iu in sku_set:
                        continue
                    if any(p in iu.lower() for p in exclude_patterns):
                        continue
                    if not any(d in iu for d in valid_cdn_patterns):
                        continue
                    seen_urls.add(iu)
                    detail_urls.append(iu)
                
                self.log(f"  分类: 主图{len(main_urls)} SKU{len(sku_items)} 详情{len(detail_urls)} (含iframe)")
                
                def find_captured(target_url):
                    if not target_url:
                        return None
                    tu = canon_img_url(target_url)
                    if tu in captured:
                        return captured[tu]
                    base = tu.split("?")[0]
                    if base in captured:
                        return captured[base]
                    tk = _url_match_key(tu)
                    if tk:
                        for cap_url, body in captured.items():
                            if _url_match_key(cap_url) == tk:
                                return body
                    try:
                        tpath = unquote(urlparse(tu).path or "")
                    except Exception:
                        tpath = ""
                    if tpath and len(tpath) > 5:
                        for cap_url, body in captured.items():
                            try:
                                if unquote(urlparse(cap_url).path or "") == tpath:
                                    return body
                            except Exception:
                                continue
                    return None
                
                async def fetch_image_bytes(abs_url):
                    """未进拦截缓存时用当前 Cookie 拉原图；带 Accept 与短重试，减少只下到部分张的情况。"""
                    u = canon_img_url(abs_url)
                    if not u or not u.startswith("http"):
                        return None
                    ref = target_url
                    try:
                        if page.url and str(page.url).startswith("http"):
                            ref = str(page.url)
                    except Exception:
                        pass
                    headers = {
                        "Referer": ref,
                        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
                    }
                    for attempt in range(3):
                        try:
                            resp = await page.request.get(u, headers=headers, timeout=45000)
                            if resp.status >= 400:
                                await page.wait_for_timeout(200 + attempt * 200)
                                continue
                            data = await resp.body()
                            if len(data) > 400:
                                return data
                        except Exception:
                            await page.wait_for_timeout(250 + attempt * 200)
                    return None
                
                for i, u in enumerate(main_urls):
                    data = find_captured(u)
                    if not data:
                        data = await fetch_image_bytes(u)
                    if data:
                        if save_jpg(data, os.path.join(save_dir, "主图", f"{i + 1}.jpg"), compress=False):
                            saved["main"] += 1
                
                seen_s = set()
                sc = 0
                for item in sku_items:
                    u = item["url"]
                    if u in seen_s:
                        continue
                    seen_s.add(u)
                    data = find_captured(u)
                    if not data:
                        data = await fetch_image_bytes(u)
                    if data:
                        sc += 1
                        name = item.get("name", "")
                        fn = f"{sc}_{name}.jpg" if name else f"{sc}.jpg"
                        fn_png = f"{sc}_{name}.png" if name else f"{sc}.png"
                        if save_jpg(data, os.path.join(save_dir, "SKU图", fn), compress=False):
                            saved["sku"] += 1
                        if matting_mode == "each":
                            dst_png = os.path.join(save_dir, "透明图", fn_png)
                            if matting_parallel:
                                async def _cutout_one(raw_bytes=data, out_path=dst_png, name=fn_png):
                                    async with cutout_sem:
                                        ok = await asyncio.to_thread(
                                            sku_cutout_bytes_to_png, raw_bytes, out_path, self.log
                                        )
                                        if not ok:
                                            self.log(f"    并发抠图失败: {name}")
                                cutout_tasks.append(asyncio.create_task(_cutout_one()))
                            else:
                                sku_cutout_bytes_to_png(data, dst_png, self.log)
                
                dc = 0
                for u in detail_urls:
                    data = find_captured(u)
                    if not data:
                        data = await fetch_image_bytes(u)
                    if data:
                        dc += 1
                        jpg_path = os.path.join(save_dir, "详情图片", f"{601 + dc}.jpg")
                        if save_jpg(data, jpg_path, compress=True, max_kb=1024):
                            saved["detail"] += 1
                
                # ====== 视频：仅保留「字节数最大」的一条有效视频（逻辑简化）=====
                self.log(f"  收集到视频链接/片段: {len(video_list)} 个")
                
                def _is_valid_video_blob(b):
                    if len(b) < 10000:
                        return False
                    if b[:4] in (b"ftyp", b"mdat", b"moov") or b"ftyp" in b[:120]:
                        return True
                    if len(b) > 200000 and b[0] == 0x47:
                        return True
                    return False
                
                seen_url = set()
                unique_videos = []
                for vu, vbody in video_list:
                    key = (vu or "").split("?")[0]
                    if key in seen_url:
                        continue
                    seen_url.add(key)
                    unique_videos.append((vu, vbody))
                self.log(f"  去重 URL 后待处理: {len(unique_videos)} 个")
                
                async def download_m3u8_video(m3u8_url, save_path):
                    """下载m3u8格式的视频"""
                    try:
                        resp = await page.request.get(m3u8_url)
                        m3u8_content = await resp.text()
                        
                        base_url = m3u8_url.rsplit("/", 1)[0]
                        ts_urls = []
                        
                        for line in m3u8_content.split("\n"):
                            line = line.strip()
                            if line and not line.startswith("#"):
                                if line.startswith("http"):
                                    ts_urls.append(line)
                                else:
                                    ts_urls.append(f"{base_url}/{line}")
                        
                        if not ts_urls:
                            self.log(f"    m3u8解析失败，无TS分段")
                            return False
                        
                        self.log(f"    解析到 {len(ts_urls)} 个TS分段，开始下载...")
                        
                        ts_data_list = []
                        for i, ts_url in enumerate(ts_urls):
                            try:
                                ts_resp = await page.request.get(ts_url)
                                ts_data = await ts_resp.body()
                                ts_data_list.append(ts_data)
                            except Exception as e:
                                self.log(f"    TS分段 {i} 下载失败: {e}")
                                continue
                            
                            if (i + 1) % 10 == 0:
                                self.log(f"    已下载 {i + 1}/{len(ts_urls)} 个TS分段...")
                        
                        if not ts_data_list:
                            self.log(f"    所有TS分段下载失败")
                            return False
                        
                        final_data = b"".join(ts_data_list)
                        
                        with open(save_path, "wb") as f:
                            f.write(final_data)
                        
                        self.log(f"    m3u8视频下载完成 ({len(final_data) / 1024 / 1024:.1f}MB)")
                        return True
                        
                    except Exception as e:
                        self.log(f"    m3u8视频下载失败: {e}")
                        return False
                
                best_blob = None
                vid_tmp = os.path.join(save_dir, "视频", "._tmp_tb_video.mp4")
                os.makedirs(os.path.dirname(vid_tmp), exist_ok=True)
                for vu, vbody in unique_videos:
                    blob = None
                    try:
                        self.log(f"    处理视频: {vu[:70]}...")
                        if vbody == b"VIDEO_PLACEHOLDER":
                            if vu.lower().endswith(".m3u8"):
                                if await download_m3u8_video(vu, vid_tmp):
                                    with open(vid_tmp, "rb") as f:
                                        blob = f.read()
                                    try:
                                        os.remove(vid_tmp)
                                    except Exception:
                                        pass
                            else:
                                try:
                                    video_resp = await page.request.get(vu, timeout=60000)
                                    blob = await video_resp.body()
                                except Exception as e:
                                    self.log(f"    ❌ MP4下载失败: {e}")
                            if blob and _is_valid_video_blob(blob):
                                if best_blob is None or len(blob) > len(best_blob):
                                    best_blob = blob
                                    self.log(f"    ✅ 当前保留最大: {len(blob) / 1024 / 1024:.2f}MB")
                            continue
                        
                        self.log(f"    视频URL: {vu[:100]}...")
                        self.log(f"    原始大小: {len(vbody) / 1024:.1f}KB")
                        
                        if vu.lower().endswith(".m3u8"):
                            try:
                                if await download_m3u8_video(vu, vid_tmp):
                                    with open(vid_tmp, "rb") as f:
                                        blob = f.read()
                                    try:
                                        os.remove(vid_tmp)
                                    except Exception:
                                        pass
                            except Exception as e:
                                self.log(f"    ❌ m3u8视频处理失败: {e}")
                                try:
                                    with open(vid_tmp.replace(".mp4", ".m3u8"), "wb") as f:
                                        f.write(vbody)
                                except Exception:
                                    pass
                        else:
                            try:
                                text_content = vbody.decode("utf-8", errors="ignore")
                                if text_content.startswith("#EXTM3U") or "#EXTINF" in text_content:
                                    self.log(f"    ⚠️ 这是M3U8播放列表文件正在尝试解析...")
                                    if await download_m3u8_video(vu, vid_tmp):
                                        with open(vid_tmp, "rb") as f:
                                            blob = f.read()
                                        try:
                                            os.remove(vid_tmp)
                                        except Exception:
                                            pass
                                    else:
                                        self.log(f"    ❌ m3u8解析下载失败")
                                    if blob and _is_valid_video_blob(blob):
                                        if best_blob is None or len(blob) > len(best_blob):
                                            best_blob = blob
                                            self.log(f"    ✅ 当前保留最大: {len(blob) / 1024 / 1024:.2f}MB")
                                    continue
                            except Exception:
                                pass
                            
                            if blob is None and len(vbody) > 50000:
                                blob = vbody
                            elif blob is None:
                                self.log(f"    ⚠️ 视频文件太小({len(vbody) / 1024:.1f}KB)，可能是无效视频")
                        
                        if blob and _is_valid_video_blob(blob):
                            if best_blob is None or len(blob) > len(best_blob):
                                best_blob = blob
                                self.log(f"    ✅ 当前保留最大: {len(blob) / 1024 / 1024:.2f}MB")
                    except Exception as e:
                        self.log(f"    视频保存失败: {e}")
                
                try:
                    if os.path.isfile(vid_tmp):
                        os.remove(vid_tmp)
                except Exception:
                    pass
                
                if best_blob:
                    stem = (product_code or "").strip() or "0"
                    outv = os.path.join(save_dir, "视频", f"{stem}.mp4")
                    os.makedirs(os.path.dirname(outv), exist_ok=True)
                    with open(outv, "wb") as f:
                        f.write(best_blob)
                    saved["video"] = 1
                    self.log(f"  已写入 1 个视频文件（取体积最大的一条），约 {len(best_blob) / 1024 / 1024:.2f}MB")
                else:
                    saved["video"] = 0
                
                if saved["video"] == 0:
                    video_dir = os.path.join(save_dir, "视频")
                    if os.path.exists(video_dir):
                        try:
                            import shutil
                            shutil.rmtree(video_dir)
                        except:
                            pass
                if cutout_tasks:
                    self.log(f"  等待并发抠图任务完成: {len(cutout_tasks)} 个")
                    await asyncio.gather(*cutout_tasks, return_exceptions=True)
                            
            except Exception as e:
                self.log(f"  出错: {e}")
            finally:
                # 复用下载页：不关标签，下一款直接 goto，避免每款弹新窗口
                try:
                    if page is not self._reuse_dl_page:
                        await page.close()
                except Exception:
                    pass
            return saved
        
        return self.call(_())
    
    def stop(self):
        if self._loop and self._loop.is_running():
            self._loop.call_soon(self._loop.stop)


class DownloadHistory:
    def __init__(self, output_dir):
        self.output_dir = output_dir
        self.history_file = os.path.join(output_dir, ".download_history.json")
        self.history = self.load()
    
    def load(self):
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, "r", encoding="utf-8") as f:
                    return json.load(f)
        except:
            pass
        return {}
    
    def save(self):
        try:
            with open(self.history_file, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
            try:
                subprocess.run(["attrib", "+h", self.history_file], capture_output=True)
            except:
                pass
        except:
            pass
    
    def is_downloaded(self, code):
        return code in self.history
    
    def mark_downloaded(self, code, result):
        self.history[code] = {
            "main": result.get("main", 0),
            "sku": result.get("sku", 0),
            "detail": result.get("detail", 0),
            "video": result.get("video", 0),
            "time": time.time(),
        }
        self.save()


import time
import json
import subprocess
