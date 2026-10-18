# 平台扫描脚本（独立）

从实训平台抓取实验手册与图片，输出为与主程序兼容的结构。

## 依赖

```bash
pip install playwright
# 使用本机 Edge（channel="msedge"），无需 playwright install
```

## 使用

1. 修改 `scan.py` 顶部配置：`BASE_URL`（平台地址）
2. 运行 `python scan.py`
3. 在弹出的浏览器中**人工登录**（图形验证码）
4. 程序检测登录成功后自动发现实验，按提示输入序号（或 `all`）抓取
5. 输出到 `lab-output/<testId>_<实验名>/`（含 `实验要求与命令.md`、`images/`、`_raw/`）

## 要点

- 使用 `ignore_https_errors=True` 应对证书缺失 SAN 的站点（免代理/CA）
- 使用 `launch_persistent_context` 持久化登录态（`.profile/`）
- 仅调用学生端接口；请勿调用教师端接口
- 抓取可能触发平台的实验会话启动行为，请自行确认可接受
