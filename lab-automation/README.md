# 远程实验环境自动化工具链 · 主程序

tkinter GUI：选择实验 → 载入/编辑步骤 → 连接服务器 → 真实终端执行并截图 → 生成 Word。

## 运行

```bash
pip install python-docx pillow paramiko
python app.py
```

首次运行会在 `%LOCALAPPDATA%\远程实验环境自动化工具链\` 生成 `config.json` 与 `steps/`。

## 配置

参考 `config.example.json`，在 GUI 的「设置」中填写：

- SSH：主机 / 端口 / 用户 / 密码（保存时 DPAPI 加密）
- gsql 端口、openGauss 数据目录
- Word 模板路径、输出目录
- 封面字段

## 步骤文件

`steps/<testId>.json`：

```json
{
  "testId": "...",
  "name": "实验名称",
  "capture_stop": false,
  "setup": ["DROP DATABASE IF EXISTS dbtrade;", "CREATE DATABASE dbtrade;", "\\c dbtrade", "..."],
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
