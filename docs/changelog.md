# 📋 更新日志

## 2026-07-27 - 账号池与项目边界调整

- 新增登录态账号池，支持筛选、权限测试、批量测试、复制账密和复制完整账号。
- 移除 `/api/external/*` 公开接口、对外 API Key 设置和验证码提取入口；项目聚焦 WebUI 邮件与账号管理。

## 2026-04-10 - 注册 / 邮箱池 API 覆盖补齐

### 新增功能
- 🧩 **注册 / 邮箱池对外接口覆盖** - 新增并统一 `/api/external/*` 下的健康检查、能力探测、账号状态、邮件摘要、最新邮件、邮件详情、原始内容、验证码、验证链接、等待新邮件、异步 probe、邮箱池领取 / 释放 / 回传 / 统计
- 📦 **邮箱池状态模型** - 账号新增 `available / claimed / used / cooldown / frozen / retired` 六种池状态，支持租约与冷却时间
- 🖥️ **页面端邮箱池面板** - 在邮箱下方面板新增池统计、当前账号池状态、接口示例、快捷切换状态
- ⚙️ **邮箱池设置项** - 设置页新增外部邮箱池接口开关、默认租约时长、冷却时长

### 数据库变更
- `accounts` 表新增：
  - `provider`
  - `pool_status`
  - `pool_claim_token`
  - `pool_claimed_by`
  - `pool_task_id`
  - `pool_claimed_at`
  - `pool_lease_expires_at`
  - `pool_cooldown_until`
  - `pool_last_result`
  - `pool_last_detail`

### API 变更
- 新增 `GET /api/external/health`
- 新增 `GET /api/external/capabilities`
- 新增 `GET /api/external/account-status`
- 新增 `GET /api/external/messages`
- 新增 `GET /api/external/messages/latest`
- 新增 `GET /api/external/messages/{message_id}`
- 新增 `GET /api/external/messages/{message_id}/raw`
- 新增 `GET /api/external/verification-code`
- 新增 `GET /api/external/verification-link`
- 新增 `GET /api/external/wait-message`
- 新增 `GET /api/external/probe/{probe_id}`
- 新增 `POST /api/external/pool/claim-random`
- 新增 `POST /api/external/pool/claim-release`
- 新增 `POST /api/external/pool/claim-complete`
- 新增 `GET /api/external/pool/stats`
- 新增 `GET /api/pool/stats`
- 新增 `POST /api/accounts/{id}/pool-state`

---

## 2026-02-24 - DuckMail 临时邮箱集成

### 新增功能
- 🦆 **DuckMail 临时邮箱** - 新增 DuckMail 作为第二个临时邮箱提供商，支持与 GPTMail 双服务商切换
  - 支持自定义 DuckMail 实例地址，默认使用官方 `https://api.duckmail.sbs`
  - 支持可选的 API Key（dk_ 前缀）用于获取私有域名
  - 创建邮箱时可选择域名、自定义用户名和密码
  - 邮件保存 3 天，账号不自动删除
- 🔄 **提供商切换** - 生成临时邮箱时可在 GPTMail 和 DuckMail 之间切换
- 🏷️ **提供商标识** - 临时邮箱列表显示 GPTMail/DuckMail 提供商标签

### 数据库变更
- `temp_emails` 表新增 `provider`、`duckmail_token`、`duckmail_account_id`、`duckmail_password` 字段

### API 变更
- 新增 `GET /api/duckmail/domains` - 获取 DuckMail 可用域名
- `POST /api/temp-emails/generate` - 新增 `provider` 参数，DuckMail 模式需传 `domain`、`username`、`password`
- `GET /api/settings` - 返回 `duckmail_base_url`、`duckmail_api_key` 字段
- `PUT /api/settings` - 支持 `duckmail_base_url`、`duckmail_api_key` 字段更新

---

## 2026-02-23 - 对外 API 支持

### 新增功能
- 🔑 **对外 API Key 管理** - 在系统设置中支持配置对外 API Key，支持随机生成
- 📡 **对外 API 端点** - 新增 `GET /api/external/emails`，通过 API Key + 邮箱 + 类型获取邮件
  - 支持 `X-API-Key` Header 或 `api_key` 查询参数认证
  - 支持 `inbox`（收件箱）和 `junkemail`（垃圾邮件）两种类型
  - 复用 Graph API → IMAP 回退逻辑
- 📚 **文档拆分** - README 拆分为模块化文档，通过外链引用

### API 变更
- 新增 `GET /api/external/emails` - 对外邮件查询 API
- `GET /api/settings` - 返回值新增 `external_api_key` 字段
- `PUT /api/settings` - 支持 `external_api_key` 字段更新

---

## 2026-02-18 - API 优先级回退与分组代理

### 新增功能
- 🌐 **分组代理** - 每个分组可配置 HTTP/SOCKS5 代理
- 🔄 **API 优先级回退** - Graph API → IMAP(新) → IMAP(旧) 自动回退
- ❌ **详细错误弹框** - 所有 API 方式均失败时，弹框展示每种方式的失败原因

---

## 2026-01-31 - 标签管理与邮件删除

### 新增功能
- 🏷️ **标签管理** - 支持给邮箱账号打标签，批量打/去标签，按标签筛选
- 🗑️ **邮件删除** - 支持单封/批量永久删除邮件，二次确认防误删

---

## 2026-01-29 - 邮件缓存与信任模式

### 新增功能
- 💾 **邮件缓存** - 智能缓存邮件列表，切换邮箱/文件夹即时展示
- 🔓 **信任邮件模式** - 支持查看原始邮件内容（绕过安全过滤）

---

## 2026-01-26 - 安全增强版本

### 新增安全特性
- 🛡️ **XSS 防护** - DOMPurify 净化 + iframe 沙箱 + 输入转义
- 🔒 **CSRF 防护** - Flask-WTF Token 验证，优雅降级
- 🔐 **敏感数据加密** - Fernet 加密 + bcrypt 哈希 + PBKDF2 密钥派生
- 🚦 **登录速率限制** - 5 次失败锁定 15 分钟
- 📋 **审计日志** - 记录导出等敏感操作
- 🔑 **二次验证** - 导出功能密码确认

### 数据库变更
- 新增 `audit_logs` 表

### API 变更
- 新增 `/api/export/verify` - 导出验证 Token
- 新增 `/api/csrf-token` - 获取 CSRF Token

---

## 2026-01-25 - Token 刷新管理增强

### 新增功能
- ⏰ **灵活的定时刷新** - 按天数或 Cron 表达式配置
- 🔔 **定时刷新开关** - 随时启用/禁用
- 📊 **刷新统计优化** - 实时显示失败邮箱数量
- 📜 **刷新历史管理** - 近半年完整历史，自动清理
- ❌ **失败邮箱标识** - 加粗标红显示
- 📝 **备注展示** - 邮箱列表显示账号备注

### 数据库变更
- 新增 `enable_scheduled_refresh`、`refresh_cron`、`use_cron_schedule` 配置

---

## 2026-01-23 - Token 刷新管理功能

### 新增功能
- ✨ **全量刷新 Token** - 一键测试所有账号
- 📊 **实时进度展示** - Server-Sent Events 流式响应
- 🔁 **失败重试** - 单个/批量重试
- 🔍 **全屏查看邮件** - 新增全屏模式

### 数据库变更
- 新增 `account_refresh_logs` 表
- `accounts` 表新增 `last_refresh_at` 字段

### API 变更
- 新增 8 个刷新相关 API 端点
