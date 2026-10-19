# 平台扫描脚本（独立）

从实训平台抓取实验手册与图片，输出为与主程序兼容的结构：

```
lab-output/<testId>_<实验名>/
    实验要求与命令.md
    images/
    _raw/studentExpInfo.json
```

## 依赖

```bash
pip install playwright
# 使用本机 Edge（channel="msedge"），无需 playwright install
```

## 使用

1. 修改 `scan.py` 顶部配置：`BASE_URL`（平台地址）
2. 运行 `python scan.py`
3. 在弹出的浏览器中**人工登录**（图形验证码）
4. 登录后自动发现实验，按提示输入序号（或 `all`）抓取
5. 输出到脚本同级的 `lab-output/`

## 鉴权要点（容易踩的坑）

- 会话凭据是 cookie（`OLE_SESSION` / `refresh_token`），鉴权头是 `Auth-Token`。
- **不要自己调用 `getCsrfToken`**：每次调用都会生成新令牌并作废旧令牌，会把浏览器里 App 的令牌顶掉，表现为“登录后闪一下就被踢回登录页”。
- 正确做法：用 `page.on("request")` 从 App 自己的请求头里**顺**一个非空 `Auth-Token`。
- 遇到 `EDU99999` 不要自己换令牌，而是重载 `/user/index` 让 App 自己换，再顺新令牌重试一次。
- 等待必须用 `page.wait_for_timeout()`；`time.sleep()` 在 Playwright 同步 API 下不会派发事件。

## 发现 / 抓取链路

```
POST /teachmgnt/api/class/user/queryClassOptionByAuth          {}
POST /teachmgnt/api/course/user/findUserCourseList             {classId}（body 建议带 pageNum/pageSize）
POST /coursemanagement/myCourse/catalogList                    {courseId, classId}
      resourceType == 3 → 实验（resourceId 即 testId）
POST /expDev/studentExpInfo/queryExpInfoById                   {expId, classId, isQueryStep: 1}
GET  /expDev/api/expFile/downloadFile?fileCode=<uuid>          → 图片
```

> 参数/字段以实际平台为准；脚本会优先让 App 自己发请求并捕获响应，失败再退回脚本构造。

## 其他要点

- `ignore_https_errors=True` 应对证书缺失 SAN 的站点（免代理/免 CA）
- `launch_persistent_context` 持久化登录态（`.profile/`，可删）
- **只调用学生端接口**；不要调用 `/expDev/expInfo/*`（教师端），否则会话会被判未登录
- **不打开 `/user/exercise`**，因此不会触发 `userStartExp`、不启动实验会话
- 请仅在你**有权访问**的平台与服务器上使用
