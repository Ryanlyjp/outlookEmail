# 接口说明

本项目当前定位为 Outlook 邮件与账号池管理工具，不提供无需登录的公开收件、验证码提取或邮箱池 API。

## 使用边界

- `/api/external/*` 对外接口已移除，访问这些路径会返回 404。
- Web 界面使用的 `/api/*` 均属于登录态内部接口，不承诺作为第三方集成 API。
- 系统设置中不再提供“对外 API Key”。
- 如后续需要程序化收件或验证码提取，建议拆分为独立服务，避免账号管理后台直接暴露敏感账号和邮件数据。

## WebUI 内部接口

账号池复用现有登录态接口：

- `GET /api/accounts`：读取账号摘要、分组和权限测试状态。
- `GET /api/accounts/<id>`：读取单个账号详情，用于编辑及复制账号信息。
- `POST /api/accounts/<id>/refresh`：对单个账号执行 Graph 与 OAuth IMAP 权限测试。
- `POST /api/accounts/refresh-selected`：批量执行权限测试。

这些接口需要有效 Web 登录会话。账号详情包含密码、Client ID 和 Refresh Token 等敏感信息，只应在受控管理环境中使用。
