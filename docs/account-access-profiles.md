# Outlook 账号访问配置

Outlook OAuth 账号通过两个独立字段决定 Token 端点和取信方式。普通密码型 IMAP 账号继续使用原有配置，不受这两个字段影响。

## 字段

### `mail_protocol`

- `auto`：先使用 Microsoft Graph，非连接类失败时回退到 OAuth IMAP。
- `graph`：只使用 Microsoft Graph，不回退 IMAP。
- `imap_oauth`：只使用 OAuth IMAP，不请求 Graph 邮件接口。

### `oauth_authority`

- `common`：使用 `https://login.microsoftonline.com/common/oauth2/v2.0/token`。
- `consumers`：使用 `https://login.microsoftonline.com/consumers/oauth2/v2.0/token`，适用于仅支持个人 Microsoft 账号的应用注册。

Graph 和 OAuth IMAP 都使用账号自身的 `oauth_authority`。协议只决定 Token scope 和后续邮件访问方式。

## 典型配置

| 号池类型 | `mail_protocol` | `oauth_authority` |
|---|---|---|
| 通用 Graph | `graph` | `common` |
| Personal account only Graph | `graph` | `consumers` |
| Personal account only OAuth IMAP | `imap_oauth` | `consumers` |
| 保持旧版自动回退行为 | `auto` | `common` |

## 数据库迁移

应用启动时会为 `accounts` 表自动增加：

```sql
mail_protocol TEXT DEFAULT 'auto'
oauth_authority TEXT DEFAULT 'common'
```

已有账号因此保持旧版 Graph 优先、IMAP 回退的行为。系统不会根据 Refresh Token 内容猜测账号类型；需要在管理界面或账号接口中明确分类。

## 接口字段

账号导入 `POST /api/accounts` 和账号更新 `PUT /api/accounts/<id>` 接受：

```json
{
  "mail_protocol": "graph",
  "oauth_authority": "consumers"
}
```

账号列表和详情接口会返回这两个字段。非法值会分别回退为 `auto` 和 `common`。

## 权限测试与测活

“权限测试”会对 Outlook OAuth 账号同时执行只读探测：

1. 使用账号配置的 `oauth_authority` 申请 Graph Token，并实际读取 `/me/messages?$top=1`。
2. 使用同一个 authority 申请 IMAP scope Token，并实际完成 XOAUTH2 登录和只读选择 `INBOX`。
3. 保存微软返回的 scope、Graph/IMAP 独立状态和综合健康状态。

综合状态包括：

- `all`：Graph 和 OAuth IMAP 都可用。
- `graph_only`：只有 Graph 实际读信成功。
- `imap_only`：只有 OAuth IMAP 实际登录成功。
- `token_invalid`：Graph 和 IMAP scope 都无法换取 Token。
- `unavailable`：可以换取部分 Token，但实际协议访问均失败。

账号是否判定为“可用”仍服从 `mail_protocol`：`graph` 必须 Graph 成功，`imap_oauth` 必须 IMAP 成功，`auto` 任一成功即可。权限测试不发送邮件，也不包含 abuse 恢复。

管理界面顶部的“号池管理”会跳转到独立路由 `/account-pool`。该页面拥有独立模板、样式和脚本，与邮件管理首页完全分离，集中显示库存统计和权限状态，并提供搜索、健康状态/邮箱后缀/分组筛选、排序、分页、全选筛选结果、单个/选中/全量测试、失败重试、失败列表、测试历史，以及复制账密或完整 OAuth 账号信息。所有操作均要求有效登录会话。

## 取信协议与过期 Token

邮件列表中的每封邮件会记录实际取信来源。Graph 列表使用 Graph 获取详情，OAuth IMAP New/Old 列表使用同一台 IMAP 服务器获取详情；自动回退后不会再把 IMAP UID 交给 Graph 查询。

转发任务只有在账号配置的全部取信路径都明确返回 Refresh Token 已过期或授权已撤销时，才会把账号标记为 `token_invalid`，并从后续定时转发扫描中排除。Graph 临时超时后回退到未授权的 IMAP scope、连接重置和代理故障不会触发该标记。邮件首页会以“Token 过期”标记这类账号；更新授权信息后需在号池管理中重新测试，测试通过后恢复转发扫描。

## 部署注意事项

部署前备份 `outlook_accounts.db`。数据库新增列不会删除或重写现有账号；如果需要回滚应用代码，可恢复旧镜像，新增列保留也不会影响旧版本读取。
