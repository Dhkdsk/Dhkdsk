# -*- coding: utf-8 -*-
"""
实验平台扫描工具（独立脚本）
—— 从 Talent 实训平台抓取实验的「手册 md + 图片」，落到与交接包一致的结构：

    <资料库>/<testId>_<实验名>/
        实验要求与命令.md
        images/
        _raw/studentExpInfo.json

主程序「实验报告自动化」把 <资料库> 当数据源，无需改动即可读取。

依赖：pip install playwright   （用 channel="msedge" 复用本机 Edge，无需下载 Chromium）

鉴权（据交接包 HANDOFF.md §3 + 实测）：
  cookie(OLE_SESSION/refresh_token) 是会话凭据；Auth-Token 头 = csrfToken。
  **关键：每次调用 getCsrfToken 都会换一个新值，旧值立刻失效。**
  所以脚本绝不自己去换令牌 —— 那样会把浏览器里 App 手里的令牌顶掉，
  App 下一个请求就会收到 EDU99999，被响应拦截器踢回登录页
  （实测现象：登录后闪一下就退出）。
  正确做法：从 App 自己的请求头里「顺」一个令牌来用（本脚本即如此）。

用法：  python scan.py

注意：本脚本只调学生端接口，且不打开 /user/exercise 页面，
      因此不会触发 userStartExp（不会启动实验会话、不消耗额度）。
"""
import json
import os
import re
import sys
import time
from datetime import datetime
from urllib.parse import urlparse

# ============================ 配置 ============================
BASE_URL = "https://<your-platform-host>"                            # 平台地址
LOGIN_PATH = "/user/login"                                           # 登录页路由（未登录时直接访问用户页会被平台判为越权）
TOKEN_PATH = "/userauth/auth/getCsrfToken"                           # 仅供了解：脚本绝不自己调它（会让 App 的令牌失效、把人踢下线）
SUCCESS_CODE = "EDU00000"                                            # 平台成功码
UNLOGIN_CODE = "EDU99999"                                            # 平台「未登录」码
LIBRARY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lab-output")   # 输出（资料库）目录
USER_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".profile")  # 持久化登录
LANG = "cn"
DEVICE_TYPE = "1"
DIAGNOSE = False                                                      # True=登录后自动跳页并抓 App 自身请求头做对比（排查用）
# 已知班级/课程（如自动发现失败可作兜底；来自交接包 lab-output/discovered-experiments.json）
FALLBACK_CLASS_ID = ""                                               # 可填已知班级 ID 作兜底
FALLBACK_COURSE_ID = ""                                             # 可填已知课程 ID 作兜底
# =============================================================

IMG_RE = re.compile(r'!\[[^\]]*\]\(/expDev/api/expFile/downloadFile\?fileCode=([0-9a-fA-F-]+)\)')


def log(*a):
    print(*a, flush=True)


def sanitize(name):
    bad = '<>:"/\\|?*'
    return ("".join("_" if c in bad else c for c in (name or "")).strip() or "unnamed")[:80]


def ensure_playwright():
    try:
        import playwright  # noqa
        return True
    except ImportError:
        log("未安装 Playwright。请先运行：\n  pip install playwright\n（本机已装 Edge，用 channel=\"msedge\" 复用，无需 playwright install）")
        return False


def keep_open(ctx):
    """出错时别急着把浏览器关掉，留给人工查看会话/页面状态。"""
    try:
        input("\n按回车关闭浏览器…")
    except Exception:
        pass
    ctx.close()


def main():
    if not ensure_playwright():
        return 2
    from playwright.sync_api import sync_playwright

    os.makedirs(LIBRARY, exist_ok=True)

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            USER_DATA_DIR,
            channel="msedge",                 # 复用本机 Edge，不下载 Chromium
            headless=False,
            ignore_https_errors=True,         # 关键：绕过"证书缺 SAN"死结
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.pages[0] if ctx.pages else ctx.new_page()
        install_sniffer(page)          # 只读请求头，顺 App 的 Auth-Token

        log("打开平台（如显示证书警告，勾选信任后继续）…")
        try:
            page.goto(BASE_URL + LOGIN_PATH, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            log("打开失败（若为证书错误，请在弹出窗口手动允许）：", e)

        log("请在弹出的浏览器里【人工登录】（图形验证码需手动输入）…")
        token = wait_login(page)
        if not token:
            log("未检测到登录成功（没从 App 的请求里顺到 Auth-Token），退出。")
            keep_open(ctx)
            return 3
        log("已登录，已从 App 请求中取得令牌：%s…" % token[:16])
        log("  等待 App 令牌稳定…")
        token = settle_token(page, quiet_ms=1500, timeout=15000)
        if not token:
            log("未拿到稳定令牌，退出。")
            keep_open(ctx)
            return 5
        log("令牌已稳定：%s…" % token[:16])
        tok = [token]

        # ---------------- 诊断：观察 App 自身请求，用于和脚本对比 ----------------
        if DIAGNOSE:
            try:
                probe_app(page, token, FALLBACK_CLASS_ID, FALLBACK_COURSE_ID)
            except Exception as ex:
                log("  诊断跳过（%r），继续抓取。" % ex)
            # 诊断里页面重载过，App 又换了令牌 —— 用最新那个
            if APP_TOKEN["v"]:
                tok[0] = APP_TOKEN["v"]

        # ---------------- 发现实验 ----------------
        exp_list = discover(page, tok)
        if not exp_list:
            log("未发现实验。")
            if APP_BODIES:
                log("  已记录的 App 请求体（排查参数用）：")
                for p, b in APP_BODIES.items():
                    log("    %s = %s" % (p, json.dumps(b, ensure_ascii=False)))
            keep_open(ctx)
            return 4
        log("\n发现 %d 个实验：" % len(exp_list))
        for i, e in enumerate(exp_list, 1):
            log("  %2d. %s  (%s)" % (i, e["title"], e["testId"]))

        log("\n输入要抓取的实验序号（逗号分隔，或 all）：")
        sel = input("> ").strip().lower()
        picks = exp_list if sel == "all" else [exp_list[int(x) - 1] for x in re.split(r"[,\s]+", sel) if x.strip().isdigit()]

        for e in picks:
            try:
                fetch_one(ctx, page, tok, e)
            except Exception as ex:
                log("  失败：%s -> %r" % (e["title"], ex))

        log("\n全部完成。输出目录：", LIBRARY)
        ctx.close()
    return 0


# 从 App 自己的请求头里「顺」到的令牌。绝不自己调 getCsrfToken（会作废 App 的令牌）。
APP_TOKEN = {"v": ""}
# 顺带记下 App 各接口的请求体，排查参数问题用（不参与正常流程）
APP_BODIES = {}


def install_sniffer(page):
    """监听页面发出的所有 API 请求，顺手把 App 用的 Auth-Token 记下来。

    只读请求头，不产生任何网络请求，因此不会干扰会话。
    """
    def on_req(req):
        try:
            if not any(k in req.url for k in PROBE_FRAGS):
                return
            at = _hdr(req).get("auth-token")
            if at:                      # 空字符串（getCsrfToken 自己那次）跳过
                APP_TOKEN["v"] = at
            if req.method == "POST" and req.post_data:
                try:
                    APP_BODIES[urlparse(req.url).path] = json.loads(req.post_data)
                except Exception:
                    pass
        except Exception:
            pass

    page.on("request", on_req)


def wait_login(page, timeout=300):
    """不主动调任何接口，只看两件事：URL 已离开登录页 + 顺到了 App 的 Auth-Token。

    用 page.wait_for_timeout 而不是 time.sleep —— 后者不会派发 Playwright 事件。
    """
    t0 = time.time()
    last = 0
    nudged = False
    while time.time() - t0 < timeout:
        page.wait_for_timeout(400)
        try:
            u = page.url
        except Exception:
            u = ""
        on_platform = u.startswith(BASE_URL)
        path = urlparse(u).path if on_platform else ""
        left_login = bool(on_platform) and "login" not in path

        if left_login and APP_TOKEN["v"]:
            return APP_TOKEN["v"]

        # 已登录但 App 迟迟没发带令牌的请求 → 主动跳一次首页，让它自己发
        if left_login and not nudged and time.time() - t0 > 8:
            nudged = True
            log("  （已离开登录页，主动跳一次首页以取得令牌）")
            try:
                page.goto(BASE_URL + "/user/index",
                          wait_until="domcontentloaded", timeout=60000)
            except Exception:
                pass

        if time.time() - last >= 10:
            last = time.time()
            log("  …等待登录中 已 %.0f 秒（当前页面 %s）" % (time.time() - t0, path or "?"))
    return None


def settle_token(page, quiet_ms=1500, timeout=15000):
    """等 App 把令牌换到稳定：连续 quiet_ms 毫秒不再变化才算稳。

    登录跳转后 App 往往还会再调一次 getCsrfToken，旧令牌立刻作废；
    不等稳定就拿去请求，就会得到 EDU99999。
    """
    t0 = time.time()
    last = APP_TOKEN["v"]
    last_change = time.time()
    while time.time() - t0 < timeout:
        page.wait_for_timeout(200)
        cur = APP_TOKEN["v"]
        if cur != last:
            last = cur
            last_change = time.time()
        elif cur and (time.time() - last_change) * 1000 >= quiet_ms:
            return cur
    return APP_TOKEN["v"] or last


def api(page, tok, path, body, quiet=False, retry=True):
    """在页面上下文里用带会话的 POST（含 Auth-Token）。每个响应都打印 code/msg 以便定位。

    tok 是单元素列表（可变）。遇到「未登录」不会自己去换令牌（那会把 App 踢掉），
    而是重新加载一次页面、让 App 自己去换，然后顺它的新令牌重试一次。
    """
    if APP_TOKEN["v"]:            # 始终用 App 最新的令牌，避免拿旧值
        tok[0] = APP_TOKEN["v"]
    status, redirected, url, ctype, txt = page.evaluate(
        "([p,b,t]) => fetch(p, {method:'POST', headers:{'Content-Type':'application/json','Auth-Token':t,'Xlang':%r},"
        " credentials:'include', body: JSON.stringify(b)})"
        ".then(async r => [r.status, r.redirected, r.url, r.headers.get('content-type')||'', await r.text()])" % LANG,
        [BASE_URL + path, body, tok[0]])
    try:
        d = json.loads(txt)
    except Exception:
        log("  非 JSON 返回 [HTTP %s%s] %s\n    Content-Type: %s\n    最终 URL: %s\n    正文前 300 字: %s"
            % (status, "，经过重定向" if redirected else "", path, ctype, url, re.sub(r"\s+", " ", txt)[:300]))
        return {}
    if not quiet:
        log("    %-56s HTTP %s  code=%s  msg=%s"
            % (path, status, d.get("code"), d.get("msg")))
    if d.get("code") == UNLOGIN_CODE:
        log("  [!] 会话未生效（%s 用户未登录）：%s" % (UNLOGIN_CODE, path))
        if retry:
            old = tok[0]
            log("  [i] 重新加载 /user/index，让 App 自己换一次令牌…")
            try:
                page.goto(BASE_URL + "/user/index",
                          wait_until="domcontentloaded", timeout=60000)
                nt = settle_token(page, quiet_ms=1200, timeout=15000)
            except Exception as ex:
                log("  [i] 重新加载失败：%s" % type(ex).__name__)
                nt = APP_TOKEN["v"]
            if nt and nt != old:
                tok[0] = nt
                log("  [i] 顺到新令牌，重试…")
                return api(page, tok, path, body, quiet=quiet, retry=False)
            log("  [i] 没顺到新令牌 → 不是令牌时效问题（会话可能真的没建立）")
    return d


# ---------------------- 诊断：观察 App 自身的请求 ----------------------
PROBE_FRAGS = ("/coursemanagement/", "/teachmgnt/", "/userauth/", "/expDev/")


def _hdr(req):
    try:
        h = dict(req.all_headers())
    except Exception:
        try:
            h = dict(req.headers)
        except Exception:
            h = {}
    return {str(k).lower(): v for k, v in h.items()}


def _dump_headers(h, indent="    "):
    for k in ("auth-token", "xlang", "content-type", "accept", "origin", "referer"):
        if k in h:
            v = h[k]
            if k == "auth-token" and len(v) > 20:
                v = v[:16] + "…（共 %d 字符）" % len(v)
            log("%s%-12s = %s" % (indent, k, v))
    if "cookie" in h:
        c = h["cookie"]
        names = sorted({p.split("=", 1)[0].strip() for p in c.split(";") if "=" in p})
        log("%s%-12s = （%d 字节；包含 %s）"
            % (indent, "cookie", len(c), ", ".join(names)))
    extra = sorted(k for k in h if k == "authorization" or k.startswith("x-"))
    if extra:
        log("%s%-12s = %s" % (indent, "其它头", ", ".join(extra)))


def probe_app(page, token, class_id, course_id):
    """让 App 自己发一次请求，抓它的「请求头 + 响应码」，用于和脚本的请求做对比。

    只做观测、不改任何状态；任何一步失败都不影响后续抓取。
    """
    log("\n=== 诊断：让 App 自己去调 /coursemanagement/ 接口 ===")
    log("（会自动跳到「课程」页与课程详情页，你不需要手动操作）")
    seen = []
    rec = {"on": True}

    def on_req(req):
        if not rec["on"]:
            return
        try:
            if any(k in req.url for k in PROBE_FRAGS):
                seen.append((req.method, req.url.split("?")[0], _hdr(req)))
        except Exception:
            pass

    page.on("request", on_req)

    def grab(frag, nav, label):
        try:
            with page.expect_response(lambda r: frag in r.url, timeout=30000) as ri:
                nav()
            resp = ri.value
            try:
                d = resp.json()
            except Exception:
                d = {}
            if not isinstance(d, dict):
                d = {}
            log("  App  %-24s -> HTTP %s  code=%s  msg=%s"
                % (label, resp.status, d.get("code"), d.get("msg")))
            return True
        except Exception as ex:
            log("  App  %-24s -> 未捕获到响应（%s）" % (label, type(ex).__name__))
            return False

    try:
        grab("queryRecentLearnCourses",
             lambda: page.goto(BASE_URL + "/user/index",
                               wait_until="domcontentloaded", timeout=60000),
             "queryRecentLearnCourses")

        grab("findUserCourseList",
             lambda: page.goto(BASE_URL + "/user/courses",
                               wait_until="domcontentloaded", timeout=60000),
             "findUserCourseList")

        ok = grab("catalogList",
                  lambda: page.goto("%s/user/courseDetail?courseId=%s&classId=%s"
                                    % (BASE_URL, course_id, class_id),
                                    wait_until="domcontentloaded", timeout=60000),
                  "catalogList")

        if not ok:
            try:
                page.goto(BASE_URL + "/user/courses",
                          wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(2000)
                # 仅诊断模式使用：把 <课程名称> 换成实际课程名
                grab("catalogList",
                     lambda: page.get_by_text("<课程名称>", exact=False)
                     .first.click(timeout=10000),
                     "catalogList(点击)")
            except Exception as ex:
                log("  （点击兜底也失败：%s）" % type(ex).__name__)
    finally:
        rec["on"] = False

    # 打印 App 实际发出的请求头（与脚本逐项对比）
    for frag in ("catalogList", "findUserCourseList", "queryRecentLearnCourses"):
        hit = [x for x in seen if frag in x[1]]
        if not hit:
            log("  App 本轮没有发出 %s" % frag)
            continue
        m, u, h = hit[-1]
        log("  App 实际请求：%s %s" % (m, u))
        _dump_headers(h)
        if "auth-token" in h:
            at = h["auth-token"]
            log("    Auth-Token 长度 %d（App 每次重载都会换新值；脚本用的是顺来的那个）" % len(at))
        else:
            log("    【App 根本没有发送 Auth-Token】")
        log("")

    log("  脚本请求头：Auth-Token=%s…, Xlang=%s, Content-Type=application/json, credentials=include"
        % (token[:8], LANG))
    if APP_BODIES:
        log("  App 请求体（顺带记录，便于核对参数）：")
        for p, b in APP_BODIES.items():
            log("    %s = %s" % (p, json.dumps(b, ensure_ascii=False)))
    log("=== 诊断结束 ===\n")


def as_list(d):
    if isinstance(d, list):
        return d
    if isinstance(d, dict):
        for k in ("content", "list", "records", "rows"):
            if isinstance(d.get(k), list):
                return d[k]
    return []


def _parse_courses(d):
    out = []
    for c in as_list((d or {}).get("data")):
        cid = c.get("courseId") or c.get("id")
        if cid:
            out.append((cid, c.get("courseName") or c.get("name") or ""))
    return out


def _parse_catalog(d):
    out = []
    data = (d or {}).get("data") or {}
    for node in (data.get("list") or []):
        for r in (node.get("resourceList") or []):
            if r.get("resourceType") == 3:
                out.append({"testId": r.get("resourceId"), "title": r.get("graphicTitle")})
    return out


def discover(page, tok):
    out = []
    # 班级
    r1 = api(page, tok, "/teachmgnt/api/class/user/queryClassOptionByAuth", {})
    if not r1 or r1.get("code") not in (None, SUCCESS_CODE):
        log("  班级接口返回 %s：%s（停止发现）" % (r1.get("code"), r1.get("msg")))
        return out
    classes = as_list(r1.get("data"))
    got_class = bool(classes)
    class_id = (classes[0].get("classId") if classes else None) or FALLBACK_CLASS_ID
    log("  班级数据 %d 条" % len(classes))
    log("classId = %s%s" % (class_id, "" if got_class else "  ← 兜底值（接口没给数据）"))

    # 课程：优先让 App 自己发（参数/令牌都由 App 保证），失败再退回脚本构造
    course_id, course_name = "", ""
    if page is not None:
        try:
            with page.expect_response(lambda r: "findUserCourseList" in r.url, timeout=30000) as ri:
                page.goto(BASE_URL + "/user/courses", wait_until="domcontentloaded", timeout=60000)
            d = ri.value.json()
            log("  App findUserCourseList -> code=%s msg=%s" % (d.get("code"), d.get("msg")))
            cs = _parse_courses(d)
            if cs:
                course_id, course_name = cs[0]
                log("  课程来源：App /user/courses（%d 条）" % len(cs))
        except Exception as ex:
            log("  App 课程页未捕获到 findUserCourseList（%s），改由脚本请求" % type(ex).__name__)

    if not course_id:
        for body in ({"classId": class_id, "pageNum": 1, "pageSize": 12},
                     {"classId": class_id}):
            r2 = api(page, tok, "/teachmgnt/api/course/user/findUserCourseList", body)
            if r2 and r2.get("code") == SUCCESS_CODE:
                cs = _parse_courses(r2)
                if cs:
                    course_id, course_name = cs[0]
                    log("  课程来源：脚本 findUserCourseList（%d 条）" % len(cs))
                    break
        if not course_id:
            r2 = api(page, tok, "/coursemanagement/stat/queryRecentLearnCourses", {"classId": class_id})
            if r2 and r2.get("code") == SUCCESS_CODE:
                cs = _parse_courses(r2)
                if cs:
                    course_id, course_name = cs[0]
                    log("  课程来源：queryRecentLearnCourses（%d 条）" % len(cs))

    got_course = bool(course_id)
    course_id = course_id or FALLBACK_COURSE_ID
    log("courseId = %s%s" % (course_id, "" if got_course else "  ← 兜底值（接口没给数据）"))

    # 目录：优先让 App 点击课程触发 catalogList
    if page is not None and course_name:
        try:
            page.get_by_text(course_name, exact=False).first.wait_for(timeout=15000)
            with page.expect_response(lambda r: "catalogList" in r.url, timeout=30000) as ri:
                page.get_by_text(course_name, exact=False).first.click(timeout=15000)
            d = ri.value.json()
            log("  App catalogList -> code=%s msg=%s" % (d.get("code"), d.get("msg")))
            out = _parse_catalog(d)
            if out:
                log("  目录来源：App 点击课程（%d 个实验）" % len(out))
        except Exception as ex:
            log("  App 目录未捕获到 catalogList（%s），改由脚本请求" % type(ex).__name__)

    if not out:
        r3 = api(page, tok, "/coursemanagement/myCourse/catalogList",
                 {"courseId": course_id, "classId": class_id})
        if not r3 or r3.get("code") not in (None, SUCCESS_CODE):
            log("  目录接口返回 %s：%s（停止发现）" % (r3.get("code"), r3.get("msg")))
            return out
        out = _parse_catalog(r3)

    for e in out:
        e["classId"] = class_id
        e["courseId"] = course_id
    return out


def fetch_one(ctx, page, tok, e):
    tid, title, class_id = e["testId"], e["title"], e["classId"]
    log("\n抓取：%s (%s)" % (title, tid))
    settle_token(page, quiet_ms=600, timeout=8000)   # 确保用的是当前有效令牌
    d = api(page, tok, "/expDev/studentExpInfo/queryExpInfoById",
            {"expId": tid, "classId": class_id, "isQueryStep": 1}).get("data") or {}

    manuals = d.get("expManualEntityList") or []
    parts = []
    for man in manuals:
        for s in man.get("expStepList") or []:
            c = (s.get("stepContent") or {}).get("mdContent")
            if c:
                parts.append(c)
    md = "\n\n".join(parts)
    if not md:
        log("  未取到手册正文；返回字段：", list(d.keys())[:10])
        return

    out_dir = os.path.join(LIBRARY, "%s_%s" % (tid, sanitize(title)))
    img_dir = os.path.join(out_dir, "images")
    raw_dir = os.path.join(out_dir, "_raw")
    os.makedirs(img_dir, exist_ok=True)
    os.makedirs(raw_dir, exist_ok=True)

    # 下载图片（引用在 md 里）
    codes = IMG_RE.findall(md)
    saved, missing = [], []
    for i, code in enumerate(dict.fromkeys(codes), 1):
        url = BASE_URL + "/expDev/api/expFile/downloadFile?fileCode=" + code
        try:
            r = ctx.request.get(url, headers={"Auth-Token": tok[0]})
            body = r.body()
            if r.ok and body:
                open(os.path.join(img_dir, code + ".png"), "wb").write(body)
                saved.append(code)
            else:
                missing.append(code)
        except Exception as ex:
            missing.append(code)
        if i % 10 == 0:
            log("    图片 %d/%d" % (i, len(codes)))
    log("  图片：成功 %d / 缺失 %d" % (len(saved), len(missing)))

    md_local = IMG_RE.sub(lambda m: "![img](images/%s.png)" % m.group(1), md)

    header = ("# %s — 实验要求与命令\n\n"
              "> **数据来源**：Talent 实训平台接口原文\n"
              "> **抓取时间**：%s\n"
              "> **实验 ID**：`%s`\n"
              "> **实验名称**：%s\n"
              "> **图片**：共 %d 张\n\n---\n\n" % (
                  title, datetime.now().strftime("%Y-%m-%d %H:%M"), tid, title, len(saved)))
    open(os.path.join(out_dir, "实验要求与命令.md"), "w", encoding="utf-8").write(header + md_local)
    json.dump({"code": "EDU00000", "data": d, "title": title},
              open(os.path.join(raw_dir, "studentExpInfo.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    log("  已保存：", out_dir)


if __name__ == "__main__":
    sys.exit(main())
