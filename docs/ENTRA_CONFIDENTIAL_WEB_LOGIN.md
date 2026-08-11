# Entra 可选登录模式与个人设备测试

本文解释原有 public client 模式与 confidential Web 模式的区别。部署管理员通过一个环境变量切换，MCP URL 和工具配置均无需改变。Web 模式可以让登录在 Conditional Access 中被识别为浏览器/confidential client 登录，但**不会绕过 Microsoft Entra 安全策略**。

```dotenv
# 原有设备/原生客户端路径
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=public

# Web SSO/BYOD 试点路径
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web
```

| 比较项 | `public` | `confidential_web` |
|---|---|---|
| Entra 平台 | Mobile and desktop applications | Web |
| 上游 client secret | 不需要 | 必须，仅保存在服务器 |
| 回调 | 原生/loopback 语义 | MCP 服务器固定 HTTPS callback |
| 常见 Conditional Access 分类 | Mobile apps and desktop clients | Browser |
| 个人设备 | 取决于对应原生客户端策略 | 取决于 Browser/BYOD 策略，不保证自动放行 |

## 1. 为什么现有部署使用两个应用注册

现有设计把两个 OAuth 角色分开：

| 应用注册 | OAuth 角色 | 当前职责 |
|---|---|---|
| MCP API App | Resource Server 与 OBO 中间层 | 暴露 `access_as_user`、定义 `mcp.*` App Roles、持有 Graph delegated permissions，并使用服务端密钥执行 OBO |
| MCP Public Client App | 上游交互式客户端 | 使用 Microsoft authorization code + PKCE 获取面向 MCP API 的用户令牌；它不能安全保存 client secret |

真正的 MCP 客户端仍然向 Bridge 动态注册。Qoder Work、WorkBuddy 或 Dify 的本地回调不需要配置到 Entra；Entra 只看到 Bridge 的固定 Microsoft callback。

## 2. Confidential Web 模式改变了什么

在 `confidential_web` 模式下，OAuth Bridge 是运行在服务器上的 confidential client：

1. 浏览器跳转至 21V Entra，使用固定 HTTPS Web Redirect URI。
2. Entra 把 authorization code 返回 MCP 服务器。
3. MCP 服务器通过后端连接把 code、PKCE verifier、client ID 和 client secret 发给 Entra token endpoint。
4. MCP 服务器保存 Microsoft token，只向 MCP 客户端返回 Bridge 自己签发的一次性 code/token。
5. 工具调用仍通过 OBO，以当前用户的 Graph delegated permissions 执行。

Client secret 不会出现在浏览器 URL 中，也不会返回给 Qoder Work、WorkBuddy、Dify 或其他 MCP 客户端。

## 3. 单应用注册模式

代码可以让 MCP API App 同时充当 confidential Web 登录客户端：

```dotenv
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web

MS_CLIENT_ID=<MCP API application client ID>
MS_CLIENT_SECRET=<MCP API secret 或证书凭据>

# 留空后自动复用上面两个值。
MS_OAUTH_CLIENT_ID=
MS_OAUTH_CLIENT_SECRET=
MS_PUBLIC_CLIENT_ID=
```

同一个 App Registration 此时承担两个角色：

- 受保护的 MCP API，包含 `access_as_user` delegated scope 和 `mcp.*` App Roles；
- confidential Web client，包含固定服务器回调和服务端凭据。

优点是 Entra 对象更少；代价是客户端与资源的信任边界合并。Web Redirect URI、暴露的 Scope、Graph 权限、App Roles 和凭据轮换仍应分别进行变更控制。

## 4. 可选的双应用 Web 模式

如果租户不接受单应用授权请求，或者 ISEC 要求客户端与资源分离，可以继续使用两个 App Registration，但把上游客户端改成 confidential Web App：

```dotenv
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web
MS_CLIENT_ID=<MCP API client ID>
MS_CLIENT_SECRET=<MCP API OBO secret>
MS_OAUTH_CLIENT_ID=<独立 Web client ID>
MS_OAUTH_CLIENT_SECRET=<该 Web client 对应的 secret>
```

Web client 必须添加 MCP API 的 `access_as_user` delegated permission 并完成管理员同意。不要把 Web client ID 与 MCP API secret 混用；显式设置 Web client ID 却没有设置匹配的 `MS_OAUTH_CLIENT_SECRET` 时，服务会 fail closed。

## 4.1 单应用还是双应用

| 方案 | 优点 | 代价 |
|---|---|---|
| 单应用 | 对象、凭据和 consent 步骤较少，搭建快 | API/resource 与登录 client 的权限、回调和凭据边界合并；轮换或误配置影响面更大 |
| 双应用 | 权责隔离；可独立轮换登录 secret 与 OBO secret；审计、回滚和 Conditional Access 定位更清楚 | 多一个 App Registration、Enterprise Application、凭据与 permission consent |

企业生产默认推荐双应用。单应用不是“不安全”，但更适合个人测试、小规模部署或已有严格变更控制的环境。应用数量不会决定个人设备能否登录；决定因素仍是客户端类型和实际命中的 Conditional Access 策略。

## 5. 在现有 MCP API App 上配置 Web 平台

使用当前已经暴露 `api://<app-id>/access_as_user` 的应用：

1. 打开 **App registrations > 21V-Graph-MCP-API > Authentication**。
2. 选择 **Add a platform > Web**。
3. 添加精确的生产回调：

   ```text
   https://<public-mcp-host>/oauth/microsoft/callback
   ```

   如果服务部署在 AgentRun 等带路径前缀的平台，必须包含完整前缀。
4. 不启用 implicit access-token 或 ID-token grant；服务使用 authorization code + PKCE。
5. 单应用 confidential Web 模式下，把 **Allow public client flows** 设置为 **No**。
6. 保持单租户，除非安全架构明确批准其他模式。
7. 在 **Certificates & secrets** 中优先使用证书；若使用 client secret，应设置较短有效期、Owner 和轮换提醒。
8. 保留现有 `access_as_user` Scope、App Roles 和已经批准的 Microsoft Graph delegated permissions。
9. 如果只允许指定人员使用，在 Enterprise Application 中保持 **Assignment required? = Yes** 并分配用户/组和 `mcp.*` App Roles。

测试期间不要删除旧 public-client 注册。切换到 confidential Web 后它不会被使用，同时可以作为快速回滚路径。

## 6. 个人设备能否登录由 Conditional Access 决定

应用类型会影响登录分类，但不会自动获得安全例外。

只有同时满足以下条件，个人设备才可能成功：

- 适用的 Conditional Access 策略允许 **Browser** 客户端类型；
- 该浏览器登录不要求 compliant/hybrid-joined device，或者策略为 BYOD 提供了获批的替代控制；
- MCP Enterprise Application 和下游 Microsoft Graph 资源没有被另一条策略继续要求合规设备；
- 用户被分配到 Enterprise Application，并拥有需要的 `mcp.*` App Roles；
- MFA、位置、登录风险等其他 Grant Controls 均满足。

如果任一适用策略覆盖所有现代客户端、覆盖 Browser，或者对 MCP/Graph 强制要求合规设备，个人电脑仍会失败。只有 Entra/ISEC 管理员可以批准策略调整，MCP 不应尝试绕过。

## 7. 推荐的 BYOD 试点方式

建议让 ISEC/Entra 管理员创建窄范围试点，而不是放松现有广域策略：

- 仅针对 MCP Enterprise Application 和小范围测试组；
- Client app type 为 Browser；
- 强制 MFA 和公司要求的登录风险控制；
- 应用公司要求的会话限制或其他 BYOD 控制；
- 排除管理员账号和高风险数据场景；
- 保留 Entra sign-in log 与 MCP audit log；
- 为试点策略设置失效/复审日期。

具体 Grant Controls 由 ISEC 决定，不属于应用代码配置。

## 8. 验证步骤

1. 使用 `MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=confidential_web` 部署。
2. 检查 `/healthz`，应包含：

   ```json
   { "microsoftLoginClientType": "confidential_web" }
   ```

3. 在 MCP 客户端中重置或删除后重新添加连接，避免继续使用旧 Bridge refresh token。
4. 生产环境保持 `MCP_OAUTH_BRIDGE_CALLBACK_DELIVERY=redirect`，让浏览器按标准 OAuth 流程直接返回客户端 loopback。`background` 可能被浏览器的 Mixed Content 或 Private Network Access 策略阻止，只适合作为可接受手动备用链接的界面优化。
5. 从个人设备发起登录。
6. 打开 **Entra ID > Monitoring & health > Sign-in logs**，找到对应事件并记录：
   - Application 与 Resource；
   - Client app；
   - Device state；
   - Conditional Access result；
   - 命中的策略名称和失败原因。
7. 该模式预期的 Client app 分类是 **Browser**，但应以 Sign-in log 为唯一事实来源。
8. 登录后依次调用 `auth_status`、`graph_get_me` 和一个低风险读取工具。若登录成功但 Graph 工具失败，通常是下游 Graph Conditional Access 或权限问题，不再是 MCP callback 问题。

## 9. 回滚

无需修改 MCP URL，只需恢复环境变量：

```dotenv
MCP_OAUTH_BRIDGE_MICROSOFT_CLIENT_TYPE=public
MS_PUBLIC_CLIENT_ID=<现有 public client ID>
MS_OAUTH_CLIENT_ID=
MS_OAUTH_CLIENT_SECRET=
```

重启服务，并在 MCP 客户端重置连接以重新获取 Bridge session。

## 10. 安全注意事项

- Confidential Web 模式提高了服务器凭据保护和轮换的重要性。
- 即使客户端是 confidential，仍继续使用 PKCE。
- Bridge token、Microsoft user token 和 Graph token 的 audience 不同，不能互相转发。
- 单应用注册不会自动解决 Bridge token 静态加密、DCR consent、集中撤销或 SIEM 接入；这些仍是独立加固项。

## 11. Microsoft 官方参考

- [Public and confidential client applications](https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-applications)
- [Conditional Access conditions - Client apps](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-conditions)
- [Conditional Access target resources and client types](https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-cloud-apps)
- [Microsoft identity platform On-Behalf-Of flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow)
