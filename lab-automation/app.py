# -*- coding: utf-8 -*-
"""实验报告自动化 - GUI (tkinter)
流程：选实验 -> 解析/编辑步骤 -> 真实终端采集 -> 生成 Word
"""
import os
import re
import sys
import json
import queue
import shutil
import subprocess
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import engine

BASE = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
        else os.path.dirname(os.path.abspath(__file__)))
# 运行数据写到可写目录（%LOCALAPPDATA%），exe 可放在任意位置（含只读目录）
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or BASE, "远程实验环境自动化工具链")
os.makedirs(DATA_DIR, exist_ok=True)
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
STEPS_DIR = os.path.join(DATA_DIR, "steps")
SHOTS_DIR = os.path.join(DATA_DIR, "shots")
# 首次运行：把随程序附带的 config/steps "种"进数据目录
if not os.path.exists(CONFIG_PATH) and os.path.exists(os.path.join(BASE, "config.json")):
    import shutil as _sh
    _sh.copyfile(os.path.join(BASE, "config.json"), CONFIG_PATH)
if os.path.isdir(os.path.join(BASE, "steps")):
    os.makedirs(STEPS_DIR, exist_ok=True)
    for _fn in os.listdir(os.path.join(BASE, "steps")):
        _dst = os.path.join(STEPS_DIR, _fn)
        if not os.path.exists(_dst):
            import shutil as _sh
            _sh.copyfile(os.path.join(BASE, "steps", _fn), _dst)

DEFAULT_CONFIG = {
    "library": "资料库",
    "template": r"模板\实验报告手册.docx",
    "output_dir": "",
    "ssh_host": "", "ssh_port": 22, "ssh_user": "",
    "ssh_password": "", "key_path": os.path.expanduser(r"~\.ssh\labshot_id"),
    "gsql_port": 5432,
    "db_data_dir": "/opt/huawei/install/data/db",
    "cols": 120, "lines": 40,
    "scanner_script": "", "scanner_python": "",
    "cover": {
        "课程名称": "", "姓名": "", "学号": "", "专业": "",
        "班级": "", "时间": "", "地点": "",
    },
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(CONFIG_PATH):
        try:
            cfg.update(json.load(open(CONFIG_PATH, encoding="utf-8")))
        except Exception:
            pass
    try:
        cfg["ssh_password"] = engine.dpapi_decrypt(cfg.get("ssh_password", ""))
    except Exception:
        pass
    return cfg


def save_config(cfg):
    out = dict(cfg)
    try:
        if out.get("ssh_password") and not str(out["ssh_password"]).startswith("enc:"):
            out["ssh_password"] = engine.dpapi_encrypt(out["ssh_password"])
    except Exception:
        pass
    json.dump(out, open(CONFIG_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=2)


PROMPT_RE = re.compile(r'^\s*(?:\[[^\]]+\]\s*)?(?:[\w.\-]+=\s*)?[#$]\s*')
IMG_RE = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')


def strip_prompt(line):
    return PROMPT_RE.sub('', line).strip()


def parse_manual(md_text):
    """把手册 md 解析成步骤草稿：[{'caption','commands':[...]}]"""
    lines = md_text.splitlines()
    cut = 0
    for i, ln in enumerate(lines):
        if ln.strip().startswith("## 实验介绍") or ln.strip().startswith("## 任务") or ln.strip().startswith("# "):
            cut = i
            break
    body = lines[cut:]
    if "## 附：命令速查" in md_text:
        body = body[:next((i for i, ln in enumerate(body) if ln.strip().startswith("## 附：命令速查")), len(body))]

    steps, pending = [], []
    caption = ""
    in_fence = False
    buf = []

    SKIP = [
        re.compile(r'^ssh\s'), re.compile(r'^su\s+-\s'), re.compile(r'^gs_om\s'),
        re.compile(r'^gsql\s+-d\s+postgres'), re.compile(r'^\[?root@'), re.compile(r'^\\?q$'),
        re.compile(r'^#'), re.compile(r'^wait\b'), re.compile(r'^--'),
    ]

    def flush():
        nonlocal buf
        if buf:
            pending.append(" ".join(buf)); buf = []

    def add_cmd(c):
        c = strip_prompt(c).strip()
        if not c or any(p.match(c) for p in SKIP):
            return
        if c.startswith("\\"):       # gsql 元命令，独立成条
            flush(); pending.append(c); return
        buf.append(c)
        if c.endswith(";"):          # SQL 语句结束，合并多行
            flush()

    def push_image():
        nonlocal pending, caption
        flush()
        cmds = [c for c in pending if c]
        if cmds:
            steps.append({"caption": caption or ("步骤 %d" % (len(steps) + 1)),
                          "commands": cmds, "wait": 2.0})
        pending = []

    for raw in body:
        s = raw.strip()
        if s.startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            add_cmd(s)
            continue
        if IMG_RE.search(raw):
            push_image()
            continue
        if re.match(r'^(?: {4,}|\t+)\S', raw):
            add_cmd(s)
            continue
        if not s:
            continue
        b = re.match(r'^\*\*(.+?)\*\*$', s)
        if b:
            caption = b.group(1).strip()
            continue
        if not s.startswith("![") and "如下所示" not in s and "可以看出" not in s and len(s) <= 40:
            if s.startswith("图") or s.startswith("步骤") or s.startswith("**"):
                caption = s.strip("*# ")
    push_image()
    return steps


GSQL_PROMPT_RE = re.compile(r'^\s*[A-Za-z_][\w.]*\s*=\s*#\s?')
SHELL_PROMPT_RE = re.compile(r'^\s*(?:\[[^\]]+\][#$]|[\w.-]+@[\w.-]+:[^#$]*[#$])\s?')
SQL_SKIP_RE = re.compile(
    r'^(?:ssh\s|su\s+-\s|gs_om\b|gsql\b|gs_[a-z]|exit\b|\\q\b|\\!|vi\b|vim\b|nano\b|'
    r'yum\b|dnf\b|apt\b|cd\b|tar\b|chown\b|chmod\b|mkdir\b|rm\b|cp\b|mv\b|cat\b|'
    r'echo\b|export\b|source\b|make\b|rpm\b|java\b|ping\b|systemctl\b|service\b|mysql\b)', re.I)
SQL_END_RE = re.compile(r';\s*(?:--.*)?$')
MYSQL_STMT_RE = re.compile(r'`|AUTO_INCREMENT|ENGINE\s*=|FOREIGN_KEY_CHECKS|SET\s+NAMES\b|utf8mb4', re.I)


def clean_manual(md_text):
    """把平台手册清洗成可自动执行的 gsql 步骤。

    规则：
      - 只保留 gsql 提示符（openGauss=# / db1=# …）下的 SQL 与 \\ 元命令；
      - Shell/OS 命令（[root@…]#、vi、yum、cd、gs_om …）一律丢弃；
      - 多行 SQL 合并成一条；相同 SQL 只保留一次；
      - 自动生成 DROP DATABASE / DROP USER 的 setup，便于重复运行。
    返回 (steps, setup)
    """
    lines = md_text.splitlines()
    start = 0
    for i, ln in enumerate(lines):
        if ln.strip().startswith("## 实验介绍"):
            start = i
            break
    body = lines[start:]
    for i, ln in enumerate(body):
        if ln.strip().startswith("## 附"):
            body = body[:i]
            break

    steps, seen_sql = [], set()
    caption, fence, mode, buf = "", False, None, []

    def finalize():
        nonlocal buf, mode, caption
        if not buf:
            mode = None
            return
        cmd = "\n".join(x.rstrip() for x in buf).strip()
        buf, mode = [], None
        if not cmd or SQL_SKIP_RE.match(cmd):
            return
        if not cmd.startswith("\\") and MYSQL_STMT_RE.search(cmd):
            return
        if cmd.startswith("\\"):
            if cmd in ("\\q",) or cmd.startswith("\\!"):
                return
            steps.append({"caption": caption or ("步骤 %d" % (len(steps) + 1)),
                          "commands": [cmd], "wait": 2.0})
            return
        norm = re.sub(r'\s+', ' ', cmd).strip()
        if not SQL_END_RE.search(norm):
            cmd += ";"
            norm += ";"
        if norm in seen_sql:
            return
        seen_sql.add(norm)
        steps.append({"caption": caption or ("步骤 %d" % (len(steps) + 1)),
                      "commands": [cmd], "wait": 2.2 if "\n" in cmd else 2.0})

    def add_meta(c):
        nonlocal caption
        c = c.strip()
        if not c or c in ("\\q",) or c.startswith("\\!") or SQL_SKIP_RE.match(c):
            return
        steps.append({"caption": caption or ("步骤 %d" % (len(steps) + 1)),
                      "commands": [c], "wait": 2.0})

    for raw in body:
        s = raw.rstrip()
        st = s.strip()
        if st.startswith("```"):
            finalize()
            fence = not fence
            continue
        if not st:
            if mode == "shell":
                mode = None
            continue
        m = re.match(r'^(#{1,6})\s+(.+)$', st)
        if m and not fence:
            finalize(); mode = None
            caption = m.group(2).strip()
            continue
        m = re.match(r'^\*\*(.+?)\*\*$', st)
        if m and not fence:
            finalize(); mode = None
            caption = m.group(1).strip()
            continue
        if re.match(r'^\s*mysql>', s):
            finalize(); mode = None
            continue
        if GSQL_PROMPT_RE.match(s):
            finalize()
            content = GSQL_PROMPT_RE.sub('', s).strip()
            if not content:
                continue
            if content.startswith("\\"):
                add_meta(content)
            else:
                buf, mode = [content], "sql"
                if SQL_END_RE.search(content):
                    finalize()
            continue
        if SHELL_PROMPT_RE.match(s):
            finalize(); mode = "shell"
            continue
        if mode == "sql":
            if fence or raw.startswith("    ") or raw.startswith("\t"):
                buf.append(st)
                if SQL_END_RE.search(st):
                    finalize()
                continue
            finalize()
        if mode == "shell":
            if raw.startswith("    ") or raw.startswith("\t"):
                continue
            mode = None
        if st.startswith("![") or st.startswith("|") or st.startswith(">"):
            continue
        if len(st) <= 30 and (st.endswith("：") or st.endswith(":")):
            caption = st.rstrip("：:。")

    finalize()

    dbs, users = [], []
    for st in steps:
        for c in st["commands"]:
            dbs += re.findall(r'\bCREATE\s+DATABASE\s+([A-Za-z_]\w*)', c, re.I)
            dbs += [m[1] for m in re.findall(
                r'\bALTER\s+DATABASE\s+([A-Za-z_]\w*)\s+RENAME\s+TO\s+([A-Za-z_]\w*)', c, re.I)]
            users += re.findall(r'\bCREATE\s+(?:USER|ROLE)\s+([A-Za-z_]\w*)', c, re.I)
    setup = []
    for n in dict.fromkeys(dbs):
        if n.lower() != "postgres":
            setup.append("DROP DATABASE IF EXISTS %s;" % n)
    for u in dict.fromkeys(users):
        setup.append("DROP USER IF EXISTS %s CASCADE;" % u)
    return steps, setup


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.experiments = []      # [{dir,name,testId,md,steps}]
        self.current = None
        self.steps = []
        self.setup = []
        self.q = queue.Queue()
        self._build()
        self.scan_library()
        self.root.after(100, self._drain)

    # ------------------------------------------------------------ UI
    def _build(self):
        self.root.title("远程实验环境自动化工具链")
        self.root.geometry("980x680")

        top = ttk.Frame(self.root, padding=6); top.pack(fill="x")
        ttk.Label(top, text="资料库:").pack(side="left")
        self.var_lib = tk.StringVar(value=self.cfg["library"])
        ttk.Entry(top, textvariable=self.var_lib, width=52).pack(side="left", padx=4)
        ttk.Button(top, text="浏览", command=self.pick_lib).pack(side="left")
        ttk.Button(top, text="重扫", command=self.scan_library).pack(side="left", padx=4)
        ttk.Button(top, text="扫描平台", command=self.do_scan_platform).pack(side="left")
        ttk.Button(top, text="环境检测", command=self.do_envcheck).pack(side="right", padx=4)
        ttk.Button(top, text="设置", command=self.open_settings).pack(side="right")

        mid = ttk.Panedwindow(self.root, orient="horizontal"); mid.pack(fill="both", expand=True, padx=6)

        lf = ttk.LabelFrame(mid, text="实验列表（可多选，Ctrl/Shift）", padding=4)
        lb_wrap = ttk.Frame(lf); lb_wrap.pack(fill="both", expand=True)
        self.lst = tk.Listbox(lb_wrap, selectmode="extended", exportselection=False, width=42)
        lb_sb = ttk.Scrollbar(lb_wrap, orient="vertical", command=self.lst.yview)
        self.lst.configure(yscrollcommand=lb_sb.set)
        self.lst.pack(side="left", fill="both", expand=True)
        lb_sb.pack(side="right", fill="y")
        self.lst.bind("<<ListboxSelect>>", self.on_select)
        ttk.Button(lf, text="删除选中实验", command=self.del_experiments).pack(fill="x", pady=(4, 0))
        mid.add(lf, weight=1)

        rf = ttk.LabelFrame(mid, text="步骤（双击编辑）", padding=4)
        cols = ("no", "caption", "cmd")
        tree_wrap = ttk.Frame(rf); tree_wrap.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", height=16)
        self.tree.heading("no", text="#"); self.tree.column("no", width=36)
        self.tree.heading("caption", text="图注"); self.tree.column("caption", width=230)
        self.tree.heading("cmd", text="命令"); self.tree.column("cmd", width=430)
        tsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self.edit_step)
        self.tree.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(int(-e.delta / 120), "units"))
        self.lst.bind("<MouseWheel>", lambda e: self.lst.yview_scroll(int(-e.delta / 120), "units"))
        btns = ttk.Frame(rf); btns.pack(fill="x", pady=4)
        for t, c in [("清洗", self.clean_current), ("全部清洗", self.clean_all),
                     ("添加", self.add_step), ("编辑", self.edit_step),
                     ("删除", self.del_step), ("上移", lambda: self.move(-1)), ("下移", lambda: self.move(1)),
                     ("保存步骤", self.save_steps)]:
            ttk.Button(btns, text=t, command=c).pack(side="left", padx=2)
        mid.add(rf, weight=2)

        bottom = ttk.Frame(self.root, padding=6); bottom.pack(fill="x")
        self.btn_init = ttk.Button(bottom, text="① 环境修复与初始化", command=self.do_init)
        self.btn_init.pack(side="left")
        self.btn_run = ttk.Button(bottom, text="② 开始生成报告", command=self.do_run)
        self.btn_run.pack(side="left", padx=6)
        self.prog = ttk.Progressbar(bottom, length=260); self.prog.pack(side="left", padx=10)

        self.log = tk.Text(self.root, height=9); self.log.pack(fill="both", padx=6, pady=(0, 6))
        logbar = ttk.Frame(self.root); logbar.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(logbar, text="清空日志", command=self.clear_log).pack(side="right")
        self._log("就绪。流程：① 扫描平台 → ② 清洗 → ③ 开始生成报告。")

        # 启动自检（仅本地，避免频繁远程连接）
        def work():
            try:
                checks = engine.check_environment(self._cfg_check(), remote=False)
                self.post(self._show_checks, checks, False)
            except Exception as e:
                self.post(self._log, "环境检测异常：" + repr(e))
        threading.Thread(target=work, daemon=True).start()

    def _log(self, msg):
        self.log.insert("end", str(msg) + "\n"); self.log.see("end")

    def clear_log(self):
        self.log.delete("1.0", "end")

    def post(self, fn, *a):
        self.q.put((fn, a))

    def _drain(self):
        try:
            while True:
                fn, a = self.q.get_nowait()
                fn(*a)
        except queue.Empty:
            pass
        self.root.after(100, self._drain)

    # ------------------------------------------------------------ config
    def pick_lib(self):
        d = filedialog.askdirectory(initialdir=self.var_lib.get())
        if d:
            self.var_lib.set(d); self.cfg["library"] = d; save_config(self.cfg); self.scan_library()

    def open_settings(self):
        w = tk.Toplevel(self.root); w.title("设置"); w.grab_set()
        fields = [("ssh_host", "SSH 主机"), ("ssh_port", "SSH 端口"), ("ssh_user", "SSH 用户"),
                  ("ssh_password", "SSH 密码"), ("gsql_port", "gsql 端口"),
                  ("db_data_dir", "openGauss 数据目录"), ("template", "Word 模板"), ("output_dir", "输出目录"),
                  ("scanner_script", "平台扫描脚本 scan.py"), ("scanner_python", "扫描用 Python")]
        ent = {}
        for i, (k, lab) in enumerate(fields):
            ttk.Label(w, text=lab).grid(row=i, column=0, sticky="e", padx=6, pady=3)
            v = tk.StringVar(value=str(self.cfg.get(k, "")))
            ttk.Entry(w, textvariable=v, width=60).grid(row=i, column=1, pady=3)
            ent[k] = v
        cover0 = self.cfg.get("cover") or {}
        ck = ["课程名称", "姓名", "学号", "专业", "班级", "时间", "地点"]
        cent = {}
        for j, k in enumerate(ck):
            ttk.Label(w, text="封面·" + k).grid(row=len(fields) + j, column=0, sticky="e", padx=6, pady=3)
            v = tk.StringVar(value=str(cover0.get(k, "")))
            ttk.Entry(w, textvariable=v, width=60).grid(row=len(fields) + j, column=1, pady=3)
            cent[k] = v
        def ok():
            for k, v in ent.items():
                val = v.get()
                if k in ("ssh_port", "gsql_port"):
                    val = int(val)
                self.cfg[k] = val
            self.cfg["cover"] = {k: cent[k].get() for k in ck}
            save_config(self.cfg); w.destroy(); self._log("设置已保存")
        ttk.Button(w, text="保存", command=ok).grid(row=len(fields) + len(ck), column=1, sticky="e", pady=8)

    # ------------------------------------------------------------ library
    def scan_library(self):
        p = self.var_lib.get()
        self.cfg["library"] = p; save_config(self.cfg)
        lib = p if os.path.isabs(p) else os.path.join(BASE, p)
        self.experiments = []
        self.lst.delete(0, "end")
        if os.path.isdir(lib):
            for name in sorted(os.listdir(lib)):
                d = os.path.join(lib, name)
                md = os.path.join(d, "实验要求与命令.md")
                if os.path.isdir(d) and os.path.exists(md):
                    tid = name.split("_")[0]
                    title = name.split("_", 1)[1] if "_" in name else name
                    self.experiments.append({"dir": d, "name": title, "testId": tid, "md": md})
                    self.lst.insert("end", title)
        # 内置步骤（随程序集成）也列为可选实验
        seen = {e["testId"] for e in self.experiments}
        if os.path.isdir(STEPS_DIR):
            for fn in sorted(os.listdir(STEPS_DIR)):
                if not fn.endswith(".json"):
                    continue
                tid = fn[:-5]
                if tid in seen:
                    continue
                try:
                    d = json.load(open(os.path.join(STEPS_DIR, fn), encoding="utf-8"))
                except Exception:
                    continue
                name = d.get("name", tid)
                self.experiments.append({"dir": None, "name": name, "testId": tid, "md": None})
                self.lst.insert("end", name)
        self._log("扫描到 %d 个实验" % len(self.experiments))

    def del_experiments(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选择要删除的实验（可 Ctrl/Shift 多选）")
            return
        exps = [self.experiments[i] for i in sel]
        names = "\n".join("· " + e["name"] for e in exps[:10])
        if len(exps) > 10:
            names += "\n… 共 %d 个" % len(exps)
        if not messagebox.askyesno("确认删除",
                                   "将删除以下实验的扫描目录和步骤文件，不可恢复：\n\n"
                                   + names + "\n\n继续？"):
            return
        lib = self.var_lib.get().strip()
        lib_abs = os.path.abspath(lib if os.path.isabs(lib) else os.path.join(BASE, lib))
        n = 0
        for exp in exps:
            tid = exp["testId"]
            d = exp.get("dir")
            if d and os.path.isdir(d):
                try:
                    if os.path.commonpath([os.path.abspath(d), lib_abs]) == lib_abs:
                        shutil.rmtree(d, ignore_errors=True)
                except ValueError:
                    pass
            for p in (os.path.join(STEPS_DIR, tid + ".json"), os.path.join(SHOTS_DIR, tid)):
                try:
                    if os.path.isdir(p):
                        shutil.rmtree(p, ignore_errors=True)
                    elif os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
            n += 1
        self.current = None
        self.steps = []
        self.setup = []
        self.refresh_tree()
        self.scan_library()
        self._log("已删除 %d 个实验" % n)

    # ------------------------------------------------------------ 平台扫描
    def scanner_script(self):
        p = str(self.cfg.get("scanner_script") or "")
        if p and os.path.exists(p):
            return p
        for cand in (os.path.join(os.path.dirname(BASE), "远程实验环境自动化工具链-扫描工具", "scan.py"),
                     os.path.join(os.path.dirname(BASE), "scanner", "scan.py"),
                     os.path.join(BASE, "scanner", "scan.py")):
            if os.path.exists(cand):
                return cand
        return ""

    def scanner_python(self):
        p = str(self.cfg.get("scanner_python") or "")
        if p and os.path.exists(p):
            return p
        if not getattr(sys, "frozen", False):
            return sys.executable
        for cand in (r"D:\python\python.exe", r"C:\Program Files\Python312\python.exe",
                     r"C:\Program Files\Python311\python.exe"):
            if os.path.exists(cand):
                return cand
        return "python"

    def set_library(self, path):
        self.var_lib.set(path)
        self.cfg["library"] = path
        save_config(self.cfg)
        self.scan_library()

    def do_scan_platform(self):
        script = self.scanner_script()
        if not script:
            messagebox.showerror("扫描平台", "找不到 scan.py。请在「设置」里填写扫描脚本路径。")
            return
        py = self.scanner_python()
        cwd = os.path.dirname(script)
        self._log("启动平台扫描：" + script)
        self._log("  在弹出的窗口里人工登录；抓取完成关闭窗口后会自动刷新资料库。")
        try:
            proc = subprocess.Popen([py, script], cwd=cwd,
                                    creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0))
        except Exception as e:
            messagebox.showerror("扫描平台", "启动失败：\n%r" % e)
            return

        def wait():
            proc.wait()
            self.post(self._log, "扫描窗口已关闭，正在刷新资料库…")
            out = os.path.join(cwd, "lab-output")
            if os.path.isdir(out):
                self.post(self.set_library, out)
            else:
                self.post(self.scan_library)
        threading.Thread(target=wait, daemon=True).start()

    # ------------------------------------------------------------ 清洗
    def clean_current(self):
        if not self.current:
            messagebox.showinfo("提示", "请先选择实验")
            return
        md = self.current.get("md")
        if not md or not os.path.exists(md):
            messagebox.showinfo("提示", "该实验没有手册 md（来自内置步骤），无法清洗。")
            return
        if self.steps and not messagebox.askyesno("确认", "当前已有步骤，清洗将覆盖，继续？"):
            return
        text = open(md, encoding="utf-8").read()
        self.steps, self.setup = clean_manual(text)
        self.refresh_tree()
        self.save_steps()
        self._log("清洗完成：%d 步，setup %d 条（已保存）" % (len(self.steps), len(self.setup)))
        if not self.steps:
            messagebox.showwarning("清洗", "没有提取到 gsql 命令。该实验可能以 Shell/OS 操作为主。")

    def clean_all(self):
        if not self.experiments:
            messagebox.showinfo("提示", "资料库为空，请先扫描平台。")
            return
        if not messagebox.askyesno("确认", "将清洗全部有手册的实验，并覆盖对应 steps 文件，继续？"):
            return
        n_ok = n_empty = 0
        for exp in self.experiments:
            md = exp.get("md")
            if not md or not os.path.exists(md):
                continue
            try:
                text = open(md, encoding="utf-8").read()
                steps, setup = clean_manual(text)
            except Exception as e:
                self._log("清洗失败：%s -> %r" % (exp["name"], e))
                continue
            if not steps:
                n_empty += 1
                self._log("  %s：无 gsql 命令（已跳过）" % exp["name"])
                continue
            os.makedirs(STEPS_DIR, exist_ok=True)
            json.dump({"testId": exp["testId"], "name": exp["name"],
                       "setup": setup, "capture_stop": True, "steps": steps},
                      open(os.path.join(STEPS_DIR, exp["testId"] + ".json"), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=2)
            n_ok += 1
            self._log("  %s：%d 步" % (exp["name"], len(steps)))
        self._log("全部清洗完成：成功 %d，跳过 %d" % (n_ok, n_empty))
        if self.current:
            self.load_steps()

    def on_select(self, _=None):
        sel = self.lst.curselection()
        if not sel:
            return
        self.current = self.experiments[sel[0]]
        self.load_steps()

    def steps_file(self):
        return os.path.join(STEPS_DIR, (self.current["testId"] if self.current else "x") + ".json")

    def load_steps(self):
        self.steps = []
        self.setup = []
        fp = self.steps_file()
        if os.path.exists(fp):
            try:
                data = json.load(open(fp, encoding="utf-8"))
                self.steps = data.get("steps", [])
                self.setup = data.get("setup", [])
                self._log("已载入已保存步骤：" + fp)
            except Exception:
                pass
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, st in enumerate(self.steps, 1):
            flat = " ; ".join(" ".join(c.split()) for c in st.get("commands", []))
            self.tree.insert("", "end", iid=str(i), values=(i, st.get("caption", ""), flat[:200]))

    def gen_draft(self):
        if not self.current:
            messagebox.showinfo("提示", "请先选择实验"); return
        if not self.current.get("md"):
            messagebox.showinfo("提示", "该实验无手册 md（来自内置步骤），可直接使用或编辑现有步骤。"); return
        if self.steps and not messagebox.askyesno("确认", "当前已有步骤，解析将覆盖现有内容，继续？"):
            return
        md = open(self.current["md"], encoding="utf-8").read()
        self.steps = parse_manual(md)
        self.refresh_tree()
        self._log("已解析手册生成 %d 个步骤草稿，请检查/编辑后保存" % len(self.steps))

    def add_step(self):
        self.steps.append({"caption": "新步骤", "commands": ["SELECT 1;"], "wait": 2.0})
        self.refresh_tree()

    def edit_step(self, _=None):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0]) - 1
        st = self.steps[i]
        w = tk.Toplevel(self.root); w.title("编辑步骤"); w.grab_set()
        ttk.Label(w, text="图注").grid(row=0, column=0, sticky="ne", padx=6, pady=4)
        cap = tk.Text(w, width=60, height=2); cap.grid(row=0, column=1, pady=4); cap.insert("1.0", st.get("caption", ""))
        ttk.Label(w, text="命令\n(多条之间用空行分隔)").grid(row=1, column=0, sticky="ne", padx=6, pady=4)
        cmd = tk.Text(w, width=60, height=10); cmd.grid(row=1, column=1, pady=4)
        cmd.insert("1.0", "\n\n".join(st.get("commands", [])))
        def ok():
            st["caption"] = cap.get("1.0", "end").strip()
            chunks = re.split(r"\n\s*\n", cmd.get("1.0", "end"))
            st["commands"] = [c.strip() for c in chunks if c.strip()]
            w.destroy(); self.refresh_tree()
        ttk.Button(w, text="确定", command=ok).grid(row=2, column=1, sticky="e", pady=6)

    def del_step(self):
        sel = self.tree.selection()
        if sel:
            del self.steps[int(sel[0]) - 1]; self.refresh_tree()

    def move(self, d):
        sel = self.tree.selection()
        if not sel:
            return
        i = int(sel[0]) - 1
        j = i + d
        if 0 <= j < len(self.steps):
            self.steps[i], self.steps[j] = self.steps[j], self.steps[i]
            self.refresh_tree()
            self.tree.selection_set(str(j + 1))

    def save_steps(self):
        if not self.current:
            return
        os.makedirs(STEPS_DIR, exist_ok=True)
        json.dump({"testId": self.current["testId"], "name": self.current["name"],
                   "setup": self.setup, "capture_stop": True, "steps": self.steps},
                  open(self.steps_file(), "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        self._log("步骤已保存：" + self.steps_file())

    # ------------------------------------------------------------ actions
    def do_envcheck(self):
        self._log("环境检测中…（含远程连通，稍候）")
        def work():
            try:
                checks = engine.check_environment(self._cfg_check(), remote=True)
            except Exception as e:
                self.post(self._log, "环境检测异常：" + repr(e)); return
            self.post(self._show_checks, checks, True)
        threading.Thread(target=work, daemon=True).start()

    def _show_checks(self, checks, popup=False):
        self._log("———— 环境检测 ————")
        bad = []
        for name, ok, detail in checks:
            self._log(("  [√] " if ok else "  [×] ") + name + "  " + str(detail))
            if not ok:
                bad.append(name)
        if popup:
            if bad:
                messagebox.showwarning("环境检测", "以下项未通过：\n\n" + "\n".join(bad) +
                                       "\n\n请在「设置」中修正后重试。")
            else:
                messagebox.showinfo("环境检测", "全部通过。")

    def _ssh(self):
        return {"host": self.cfg["ssh_host"], "port": int(self.cfg["ssh_port"]),
                "user": self.cfg["ssh_user"], "key_path": self.cfg["key_path"]}

    def tpl(self):
        t = self.cfg.get("template", "")
        return t if os.path.isabs(t) else os.path.join(BASE, t)

    def _cfg_check(self):
        c = dict(self.cfg); c["template"] = self.tpl(); return c

    def do_init(self):
        self.btn_init.config(state="disabled")
        threading.Thread(target=self._init_worker, daemon=True).start()

    def _init_worker(self):
        L = lambda m: self.post(self._log, m)
        try:
            L("环境修复/初始化：检查本机 OpenSSH…")
            engine.enable_windows_openssh(log=L)
            ssh = self._ssh()

            L("安装免密密钥…")
            try:
                engine.ensure_ssh_key(self.cfg["ssh_host"], int(self.cfg["ssh_port"]), self.cfg["ssh_user"],
                                      self.cfg["ssh_password"], self.cfg["key_path"], log=L)
            except Exception as e:
                L("安装密钥失败（若已有可用密钥可忽略）：" + repr(e))

            L("探测远端 openGauss 环境…")
            env = engine.detect_remote_env(ssh, self.cfg.get("ssh_password") or None, log=L)
            if env.get("port"):
                self.cfg["gsql_port"] = int(env["port"])
            if env.get("data_dir"):
                self.cfg["db_data_dir"] = env["data_dir"]

            if not env.get("installed"):
                L("未检测到 openGauss（未找到 gsql）。请在服务器安装：openGauss 官方安装包 + simpleInstall，"
                  "或把 gsql 加入 PATH 后重试。")
            else:
                L("设置数据库本地免密…")
                try:
                    engine.enable_local_trust(self.cfg["ssh_host"], int(self.cfg["ssh_port"]), self.cfg["ssh_user"],
                                              self.cfg["ssh_password"], self.cfg["db_data_dir"], log=L)
                except Exception as e:
                    L("设置数据库免密失败：" + repr(e))
                engine.ensure_remote_running(ssh, self.cfg.get("ssh_password") or None,
                                             self.cfg["db_data_dir"], log=L)
            self.post(save_config, self.cfg)
            L("初始化完成 √")
        except Exception as e:
            L("初始化失败：" + repr(e)); L(traceback.format_exc())
        finally:
            self.post(lambda: self.btn_init.config(state="normal"))

    def do_run(self):
        sel = self.lst.curselection()
        if not sel:
            messagebox.showinfo("提示", "请选择实验"); return
        self.btn_run.config(state="disabled")
        threading.Thread(target=self._run_worker, args=(list(sel),), daemon=True).start()

    def _load_steps_for(self, exp):
        fp = os.path.join(STEPS_DIR, exp["testId"] + ".json")
        if os.path.exists(fp):
            try:
                d = json.load(open(fp, encoding="utf-8"))
                return d.get("steps", []), d.get("setup", []), d.get("capture_stop", True)
            except Exception:
                pass
        return [], [], True

    def _run_worker(self, idxs):
        try:
            for n, idx in enumerate(idxs, 1):
                exp = self.experiments[idx]
                steps, setup, capture_stop = self._load_steps_for(exp)
                if not steps:
                    self.post(self._log, "[%d/%d] 跳过（无步骤）：%s" % (n, len(idxs), exp["name"]))
                    continue
                self.post(self._log, "[%d/%d] 开始采集：%s" % (n, len(idxs), exp["name"]))
                workdir = os.path.join(SHOTS_DIR, exp["testId"])
                os.makedirs(workdir, exist_ok=True)
                engine.ensure_remote_running(self._ssh(), self.cfg.get("ssh_password") or None,
                                             self.cfg.get("db_data_dir"), log=lambda m: self.post(self._log, m))
                items = engine.run_capture(
                    self._ssh(), int(self.cfg["gsql_port"]), steps, workdir, setup=setup,
                    capture_stop=capture_stop,
                    log=lambda m: self.post(self._log, m),
                    progress=lambda p: self.post(self.prog.config, {"value": p}),
                    db_data_dir=self.cfg["db_data_dir"], cols=int(self.cfg["cols"]), lines=int(self.cfg["lines"]))
                purpose = "通过本次实验，熟悉并掌握 openGauss 数据库的基本操作方法，加深对相关概念的理解。"
                device = "已安装 openGauss 的 Linux 服务器一台。通过 SSH 登录 omm 用户，使用 gsql 连接数据库（端口 %s），按实验步骤执行并观察结果。" % self.cfg["gsql_port"]
                result = "按实验步骤依次执行各命令，结果均符合预期，具体命令与运行结果见 1.3 的截图。"
                problem = "实验过程中个别命令首次执行报错，根据提示调整后重新执行即成功。"
                feeling = "通过本次“%s”实验，掌握了 openGauss 的相关操作方法，提升了动手实践能力。" % exp["name"]
                out = os.path.join(self.cfg["output_dir"], "%s_已填写.docx" % exp["name"].split(" ")[0])
                engine.build_docx(self.tpl(), out, exp["name"], purpose, device, result, problem, feeling,
                                  items, cover=self.cfg.get("cover"), log=lambda m: self.post(self._log, m))
                self.post(self._log, "[%d/%d] 完成：%s" % (n, len(idxs), out))
            self.post(self._log, "全部完成 √")
            self.post(messagebox.showinfo, "完成", "报告已生成到：\n" + self.cfg["output_dir"])
        except Exception as e:
            self.post(self._log, "失败：" + repr(e))
            self.post(self._log, traceback.format_exc())
        finally:
            self.post(lambda: self.btn_run.config(state="normal"))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
