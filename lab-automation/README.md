# 远程实验环境自动化工具链 · 主程序

一个 tkinter GUI 串起三步：**扫描平台 → 清洗手册 → 生成报告**。

## 流程

1. **扫描平台**
   调用 `scanner/scan.py`（Playwright + 本机 Edge），人工登录后抓取实验手册与图片到 `lab-output/`。
   关闭扫描窗口后，资料库自动刷新。

2. **清洗**
   把手册 md 清洗成可直接执行的 gsql 步骤：
   - 只保留 gsql 提示符（`openGauss=#`、`db1=#` …）下的 SQL 与 `\` 元命令；
   - 丢弃 Shell/OS 命令（`vi`、`yum`、`cd`、`gs_om`、`tar` …）；
   - 多行 SQL 合并为一条，相同 SQL 只保留一次；
   - MySQL 语法（反引号、`AUTO_INCREMENT`、`utf8mb4` …）自动跳过；
   - 自动生成 `DROP DATABASE / DROP USER` 的 `setup`，便于重复运行；
   - 结果写入 `steps/<testId>.json`，可在右侧表格复核、编辑。
   - 纯 OS/安装类实验清洗后为 0 步，会被跳过（属正常）。

3. **生成报告**
   SSH 连接服务器 → 真实终端逐条执行 → 截图 → 按 Word 模板生成报告（封面、章节、图注）。

## 运行

```bash
pip install python-docx pillow paramiko
pip install playwright        # 仅「扫描平台」需要
python app.py
```

首次运行会在 `%LOCALAPPDATA%\远程实验环境自动化工具链\` 生成 `config.json` 与 `steps/`。

## 界面按钮

- **扫描平台**：启动 `scan.py`；**重扫**：重新读取资料库；**浏览**：切换资料库
- **清洗 / 全部清洗**：生成或覆盖步骤；**保存步骤**：保存当前表格
- **删除选中实验**：删除 `lab-output` 目录、`steps`、`shots`（会二次确认）
- **清空日志**：清空下方日志区
- **① 环境修复与初始化** / **② 开始生成报告**

## 配置

参考 `config.example.json`，在 GUI 的「设置」中填写：

- SSH：主机 / 端口 / 用户 / 密码（保存时 DPAPI 加密）
- gsql 端口、openGauss 数据目录
- Word 模板路径、输出目录
- 平台扫描脚本 `scan.py`、扫描用 Python
- 封面字段

## 步骤文件

`steps/<testId>.json`：

```json
{
  "testId": "...",
  "name": "实验名称",
  "capture_stop": true,
  "setup": ["DROP DATABASE IF EXISTS dbtrade;", "DROP USER IF EXISTS demo CASCADE;"],
  "steps": [
    {"caption": "图注", "commands": ["SQL 或 gsql 元命令"], "wait": 2.0}
  ]
}
```

- `setup`：连接后、截图前执行的准备/清理命令（不截图）
- `commands`：按顺序输入终端；跨行 SQL 用 `\n` 保留
- `capture_stop`：结尾是否额外截取“退出并停库”的图

## 依赖说明

- 需要本机有 `ssh` / `ssh-keygen`（Windows 可选功能里的 OpenSSH 客户端）
- 截图依赖 Windows 控制台 API（Windows 专用）
- Word 生成不依赖 Office 安装
