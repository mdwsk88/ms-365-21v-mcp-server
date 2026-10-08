# 快速上手：先成功读取自己的资料

[返回首页](../README.md) · [完整部署参考](../DEPLOYMENT.md)

目标是验证一次完整的“客户端登录 → MCP → Graph /me”调用，不是一开始就开放全部工具。需要管理员在 21V Entra 中完成配置；脚本不能代替授权、管理员同意或设备合规检查。

已有管理员提供 MCP 地址的普通用户，直接阅读[用户使用说明](../USER_GUIDE.md)，不需要执行下面的部署步骤。

## 1. 准备账号与运行环境

准备一个 Microsoft 365 operated by 21Vianet 租户、可创建应用并授予管理员同意的管理员，以及 Node.js 22+、npm 和 Git。

本机试用采用 `http://localhost:3000`，客户端、浏览器和服务需在能访问这个地址的同一台电脑上。远程服务器或云端客户端采用真实 HTTPS 地址。**云端 Dify 等平台无法通过 localhost 访问你的电脑。**

以下地址必须对应同一部署：

| 用途 | 本机示例 |
|---|---|
| 服务基地址，不含 `/mcp` | `http://localhost:3000` |
| 客户端填写的 MCP 地址 | `http://localhost:3000/mcp` |
| Entra Web 回调 | `http://localhost:3000/oauth/microsoft/callback` |

远程部署将上述基地址统一换成实际 HTTPS 地址。有平台路径前缀时，例如 `/tools/service`，也要一致保留。

## 2. 在 21V Entra 中配置一条应用注册

打开 [Azure 中国门户](https://portal.azure.cn)，进入 Microsoft Entra ID → 应用注册。

1. 创建单租户应用，记录 Tenant ID 和 Application (client) ID。
2. 在“证书和密码”创建 client secret，安全保存它的 **Value**，不要把 Secret ID 当成密码。
3. 在“公开 API”设置 Application ID URI 为 `api://<API_CLIENT_ID>`，添加 delegated scope `access_as_user`，采用管理员同意策略。
4. 在“身份验证”添加 **Web** 平台，注册上表中的精确回调。不要启用 implicit grant 或 public client flow。
5. 在“API 权限”添加 Microsoft Graph **委托权限** `User.Read`，执行管理员同意。不要添加 Application permissions 来替代它。
6. 将 [entra-app-roles.json](entra-app-roles.json) 中需要的 `appRoles` 合并到应用 Manifest；不要用该片段覆盖整个 Manifest。然后在“企业应用程序 → 此应用 → 用户和组”给测试用户分配 `mcp.users`。

此流程使用仓库已有的单应用 confidential Web 登录模式，API 和登录共用 client ID / secret。双应用隔离、租户特别的预授权/同意要求和设备策略见 [Entra 登录说明](ENTRA_CONFIDENTIAL_WEB_LOGIN.md)。

## 3. 生成最小配置

```bash
git clone https://github.com/mdwsk88/ms-365-21v-mcp-server.git
cd ms-365-21v-mcp-server
npm run setup
```

`setup` 只用 Node 内置模块，不依赖先运行 `npm ci`。输入 Tenant ID、API Client ID 和服务基地址；本机可接受默认值。它不会访问微软或修改应用注册。

编辑生成的 `.env`，将 `MS_CLIENT_SECRET` 的占位符替换成真实 secret Value。包含 `#` 或空格的值使用引号。不要把 `.env`、secret 或 token 发进聊天和 Issue。

生成器会保持以下设置：

```dotenv
MCP_OAUTH_BRIDGE_ENABLED=true
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web
MCP_INBOUND_AUTH_DISABLED=false
MCP_ALLOW_UNAUTHENTICATED_DISCOVERY=false
MCP_ROLE_BASED_FILTERING=true
MCP_TOOL_CATEGORIES=users
MCP_DISABLED_GRAPH_SCOPES=User.Read.All,User.ReadBasic.All
MCP_SEND_MODE=confirm
MCP_CONFIRM_OPERATIONS=all
MCP_AUDIT_LOG_ENABLED=true
```

它还会生成随机管理令牌，不在终端打印令牌值；在支持 POSIX 权限的平台上以 `0600` 创建文件。Windows 上仍需使用合适的文件 ACL。已有 `.env` 时拒绝覆盖，请人工维护现有部署。

中国区 Graph 使用 `https://microsoftgraph.chinacloudapi.cn/v1.0`。Microsoft 文档分别列出 `login.chinacloudapi.cn` 和 `login.partner.microsoftonline.cn` 中国区身份端点；本向导保留仓库的 partner 默认值。以管理员确认的租户 authority 为准，不要切换成全球云 `login.microsoftonline.com`。参见 [Graph 国家云](https://learn.microsoft.com/zh-cn/graph/deployments)与 [Entra 国家云认证](https://learn.microsoft.com/en-us/entra/identity-platform/authentication-national-cloud)。

### 从完整模板迁移时特别注意

`MCP_RESOURCE_URL` 是显式覆盖项。只修改 `MCP_PUBLIC_BASE_URL`，却保留旧 `.env` 中的 `MCP_RESOURCE_URL=https://mcp.example.cn/mcp`，会导致客户端发现错误的资源地址。

一般应清空这个覆盖项，让服务从基地址自动推导；自定义代理场景才显式填写并逐一验证。向导生成的配置不会添加该覆盖项，也不需要 `MS_PUBLIC_CLIENT_ID` 或第二组 Web client 凭据。

## 4. 检查并启动

```bash
npm run doctor
npm ci
npm run build
npm run start:http
```

`doctor` 有错误时退出码为 1，修正后再启动。只有警告时退出码仍为 0，应逐项确认含义；这不是生产安全验收。

打开另一个终端检查服务：

```bash
curl http://127.0.0.1:3000/healthz
```

Windows PowerShell 可使用 `curl.exe` 或 `Invoke-RestMethod`。Docker 用户应转到[部署参考](../DEPLOYMENT.md)，并确认 `.tokens` 挂载目录允许容器的非 root 用户写入；不建议用 `chmod 777` 解决权限问题。

## 5. 验证客户端与真实 Graph 调用

在支持 Streamable HTTP 和 OAuth 的客户端添加 `/mcp` 地址，完成浏览器登录。先请求“查看我的认证状态”，再请求“获取我的个人资料”。对应工具是 `auth_status` 和 `graph_get_me`。

验收时区分三层结果：服务 `/healthz` 可达，只证明进程可响应；能看到工具，只证明发现链路可用；成功返回自己的 Graph 资料，才证明本场景的登录、角色和 Graph 授权链路跑通。

未带 Bearer token 访问受保护的 `/mcp` 返回 `401` 和 OAuth metadata 提示是正常现象。不要为了让它返回 `200` 而关闭认证。每次变更角色或权限后，重新登录并按需刷新客户端工具列表。

## 启用第一个业务场景：只读邮件

先让管理员在同一 MCP API 应用中增加 Graph delegated `Mail.Read` 并同意，给测试用户增加 `mcp.mail`，保留 `mcp.users`。

对于本向导新生成的配置，将下面两项改为：

```dotenv
MCP_TOOL_CATEGORIES=users,mail
MCP_DISABLED_GRAPH_SCOPES=User.Read.All,User.ReadBasic.All,Mail.ReadWrite,Mail.Send
```

已有企业配置应**合并**禁用列表，不能覆盖掉其他尚未获批的 scope。重启服务并重新登录后，尝试“列出最近 5 封邮件，只显示主题和发件人”。此配置隐藏需要 `Mail.ReadWrite` 或 `Mail.Send` 的写工具。

其他模块按[工具目录](TOOL_CATALOG.md)逐项添加。不要只删除禁用项而忘记租户授权，不要通过关闭角色过滤来代替给用户分配角色。

## 遇到问题时

```bash
npm run doctor
npm run --silent doctor -- --json
npm run doctor -- --config /path/to/config.env
```

最后一条仅选择要检查的文件，不会改变服务实际读取的 `.env`。继承的进程环境优先于文件值，排障时也要检查部署平台的环境变量。

| 现象 | 优先检查 |
|---|---|
| `setup` 提示 `.env` 已存在 | 按预期保护原配置；人工编辑，不要直接删除生产配置 |
| `MS_CLIENT_SECRET` 错误 | 未替换占位符、误填 Secret ID 或未配置凭据 |
| `MCP_RESOURCE_URL_OVERRIDE` | 显式资源地址覆盖了服务基地址，检查旧模板与代理路径 |
| `AUTH_DISABLED` | 恢复入站认证；不要以关闭认证作为快速开始路径 |
| 只有基础工具，没有个人资料工具 | `mcp.users` 是否分配到正确的 API 企业应用，登录是否已刷新 |
| 只有个人资料，没有邮件 | 初始配置有意只加载 users；按只读邮件步骤添加权限和模块 |
| 回调不匹配、反复登录 | Entra 回调、基地址、路径前缀、secret 是否一致 |
| MFA / Conditional Access / 设备合规拒绝 | 由租户管理员查看登录日志；该工具不能绕过组织策略 |
| Graph 返回 403 | 检查 delegated consent、用户数据权限及对应 scope；离线检查无法确认它们 |

`doctor` 不连接网络、不输出配置值、不修改云端或配置文件。它针对常见 HTTP/OAuth 配置，并不是完整配置解释器、网络探测器或漏洞扫描器；特殊代理、外部 IdP、stdio 和其他高级部署需要单独验证。

提交问题前，阅读[安全策略](../SECURITY.md)，使用仓库的问题模板，仅提供经过检查的脱敏诊断信息。
