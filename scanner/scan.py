# -*- coding: utf-8 -*-
"""
实验平台扫描工具（独立脚本）
—— 从 Talent 实训平台抓取实验的「手册 md + 图片」，落到与交接包一致的结构：

    <资料库>/<testId>_<实验名>/
        实验要求与命令.md
        images/
        _raw/studentExpInfo.json

主程序「实验报告自动化」把 <资料库> 当数据源，无需改动即可读取。

依赖：pip install playwright  &&  playwright install msedge   (或本机已装 Edge)

用法：  python scan.py
"""
import json
import os
import re
import sys
import time
from datetime import datetime

# ============================ 配置 ============================
BASE_URL = "https://<your-platform-host>"                       # 平台地址
LIBRARY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lab-output")   # 输出（资料库）目录
USER_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".profile")  # 持久化登录
LANG = "cn"
DEVICE_TYPE = "1"
# 已知班级/课程（如自动发现失败可作兜底；来自交接包 discovered-experiments.json）
FALLBACK_CLASS_ID = ""
FALLBACK_COURSE_ID = ""
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
        log("未安装 Playwright。请先运行：\n  pip install playwright\n  playwright install msedge")
        return False


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

        log("打开平台（如显示证书警告，勾选信任后继续）…")
        try:
            page.goto(BASE_URL + "/user/index", wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            log("打开失败（若为证书错误，请在弹出窗口手动允许）：", e)

        log("请在弹出的浏览器里【人工登录】（图形验证码需手动输入）…")
        if not wait_login(page):
            log("未检测到登录成功，退出。")
            ctx.close()
            return 3
        log("已登录。")
        time.sleep(1.5)

        token = page.evaluate("() => sessionStorage.getItem('authToken') || (window.db && db.authToken) || ''")
        log("authToken:", (token[:16] + "…") if token else "(空，可能仅靠 cookie)")

        # ---------------- 发现实验 ----------------
        exp_list = discover(page, token)
        if not exp_list:
            log("未发现实验。")
            ctx.close()
            return 4
        log("\n发现 %d 个实验：" % len(exp_list))
        for i, e in enumerate(exp_list, 1):
            log("  %2d. %s  (%s)" % (i, e["title"], e["testId"]))

        log("\n输入要抓取的实验序号（逗号分隔，或 all）：")
        sel = input("> ").strip().lower()
        picks = exp_list if sel == "all" else [exp_list[int(x) - 1] for x in re.split(r"[,\s]+", sel) if x.strip().isdigit()]

        for e in picks:
            try:
                fetch_one(ctx, page, token, e)
            except Exception as ex:
                log("  失败：%s -> %r" % (e["title"], ex))

        log("\n全部完成。输出目录：", LIBRARY)
        ctx.close()
    return 0


def wait_login(page, timeout=300):
    """轮询侦测登录成功：路径回到 /user 且页面无密码框。每 10 秒提示一次。"""
    t0 = time.time()
    last = 0
    while time.time() - t0 < timeout:
        try:
            ok = page.evaluate(
                "() => location.pathname.startsWith('/user') && !document.querySelector('input[type=password]')")
            if ok:
                return True
        except Exception:
            pass
        if time.time() - last >= 10:
            last = time.time()
            log("  …等待登录中（请在浏览器里完成登录）已 %.0f 秒" % (time.time() - t0))
        time.sleep(2)
    return False


def api(page, token, path, body):
    """在页面上下文里用带会话的 POST（含 Auth-Token）。"""
    txt = page.evaluate(
        "([p,b,t]) => fetch(p, {method:'POST', headers:{'Content-Type':'application/json','Auth-Token':t,'Xlang':%r},"
        " credentials:'include', body: JSON.stringify(b)}).then(r => r.text())" % LANG,
        [BASE_URL + path, body, token])
    try:
        return json.loads(txt)
    except Exception:
        log("  非 JSON 返回：", txt[:200])
        return {}


def as_list(d):
    if isinstance(d, list):
        return d
    if isinstance(d, dict):
        for k in ("content", "list", "records", "rows"):
            if isinstance(d.get(k), list):
                return d[k]
    return []


def discover(page, token):
    out = []
    # 班级
    classes = as_list(api(page, token, "/teachmgnt/api/class/user/queryClassOptionByAuth", {}).get("data"))
    class_id = (classes[0].get("classId") if classes else None) or FALLBACK_CLASS_ID
    log("classId =", class_id)
    # 课程
    courses = as_list(api(page, token, "/teachmgnt/api/course/user/findUserCourseList", {"classId": class_id}).get("data"))
    if not courses:
        courses = as_list(api(page, token, "/teachmgnt/api/stat/queryRecentLearnCourses", {"classId": class_id}).get("data"))
    course_id = (courses[0].get("courseId") if courses else None) or FALLBACK_COURSE_ID
    log("courseId =", course_id)
    # 目录树
    cat = api(page, token, "/coursemanagement/myCourse/catalogList",
              {"courseId": course_id, "classId": class_id}).get("data") or {}
    for node in (cat.get("list") or []):
        for r in (node.get("resourceList") or []):
            if r.get("resourceType") == 3:
                out.append({"testId": r.get("resourceId"), "title": r.get("graphicTitle"),
                            "classId": class_id, "courseId": course_id})
    return out


def fetch_one(ctx, page, token, e):
    tid, title, class_id = e["testId"], e["title"], e["classId"]
    log("\n抓取：%s (%s)" % (title, tid))
    d = api(page, token, "/expDev/studentExpInfo/queryExpInfoById",
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
            r = ctx.request.get(url, headers={"Auth-Token": token})
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
