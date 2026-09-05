# MS 365-21V MCP Server

**让支持 MCP 的 AI 客户端，以你的身份使用世纪互联 Microsoft 365。**

查询邮件、安排日程、检索 OneDrive 和 SharePoint，连接 Microsoft Graph 中国区；保留用户权限、操作确认和审计边界。

[![CI](https://github.com/mdwsk88/ms-365-21v-mcp-server/actions/workflows/ci.yml/badge.svg)](https://github.com/mdwsk88/ms-365-21v-mcp-server/actions/workflows/ci.yml)
[![License: Apache-2.0](https://img.shields.io/badge/License-Apache--2.0-blue.svg)](LICENSE)

[快速上手](docs/QUICKSTART.md) · [完整部署](DEPLOYMENT.md) · [使用场景](USER_GUIDE.md) · [工具目录](docs/TOOL_CATALOG.md) · [English](README.en.md)

> 非 Microsoft 或 21Vianet 官方产品。需要世纪互联租户和管理员授权，不适用于直接连接全球版 Microsoft 365 或个人 Outlook 账号。本文中的业务示例需要对应模块与权限，初始配置只开放个人资料查询。

## 先看它能帮你做什么

| 你对 AI 说 | 能力 | 启用条件 |
|---|---|---|
| “获取我的个人资料，确认当前登录账号。” | 验证 OAuth → Graph 链路 | 快速上手的默认场景 |
| “列出最近 5 封邮件，只显示主题和发件人。” | 邮件读取 | Mail 模块、`Mail.Read`、`mcp.mail` |
| “检查明天的会议是否有时间冲突。” | 日程冲突分析 | Calendar + Smart 模块、对应权限与角色 |
| “在我的 OneDrive 中查找项目计划。” | 文件检索 | OneDrive 模块、`Files.Read`、`mcp.drive` |
| “整理回复内容，发送前让我确认。” | 邮件写入与确认 | 额外写权限；保持 `MCP_SEND_MODE=confirm` |

工具会返回数据供 AI 客户端使用，不会因为接入 MCP 而授予用户原本没有的数据权限。完整角色、scope 和工具映射见[工具目录](docs/TOOL_CATALOG.md)。

## 为什么专门做一个 21V 版本

- **中国区身份与数据端点**：围绕 21V Entra 和 Microsoft Graph 中国区设计，而不是只替换全球版服务的域名。
- **客户端接入更统一**：Streamable HTTP + OAuth bridge，通过一个 MCP 地址连接；服务端使用当前用户的 delegated permissions / OBO。
- **可控制的工具与写操作**：App Roles、按模块与 scope 过滤、常用工具直达/长尾工具发现，以及确认与脱敏审计。

覆盖邮件、日历、OneDrive、SharePoint、Teams、联系人、组织用户、Microsoft Search 和智能聚合。**代码中实现了工具，不等于你的租户已授权，也不等于每个全球云 API 都在 21V 可用。**

## 快速开始

### 只使用别人已经部署好的服务？

不需要克隆仓库，也不需要自己创建 Entra 应用。向管理员获取 HTTPS MCP 地址及使用权限，在支持 Streamable HTTP 和 OAuth 的客户端添加它，然后登录。参见[用户使用说明](USER_GUIDE.md)。

### 第一次部署？从“读取我的资料”开始

准备 Node.js 22+、Git，以及能配置 Entra 应用、同意权限并分配角色的管理员。先按[快速上手](docs/QUICKSTART.md)完成单应用配置，再运行：

```bash
git clone https://github.com/mdwsk88/ms-365-21v-mcp-server.git
cd ms-365-21v-mcp-server
npm run setup
```

向导只询问 Tenant ID、API Client ID 和服务地址。它生成最小 `.env`，不会覆盖已有文件，也不会要求你把 client secret 放进命令行。编辑 `.env` 中的 `MS_CLIENT_SECRET`，使用 secret 的 **Value**，不是 Secret ID，然后运行：

```bash
npm run doctor
npm ci
npm run build
npm run start:http
```

本机桌面客户端连接 `http://localhost:3000/mcp`。远程或云端客户端需要能够访问的 HTTPS 地址；云端客户端的 `localhost` 不是你的电脑。

**成功标准：** 在客户端完成登录，调用 `auth_status`，再成功执行 `graph_get_me`。健康检查通过或能看到工具，都不等于 Graph 权限已经配置成功。

初始配置保持 OAuth、App Roles、审计和操作确认开启，只加载个人资料模块。后续按需启用邮件、日历等功能，不用一次申请全部权限。[逐步启用只读邮件](docs/QUICKSTART.md#启用第一个业务场景只读邮件)

已有部署请先运行 `npm run doctor`，不要重新生成或覆盖 `.env`。Docker、双应用、反向代理和生产检查见[完整部署说明](DEPLOYMENT.md)。

## 配置错了，从哪里查？

```bash
npm run doctor
npm run --silent doctor -- --json
```

离线检查会提示缺失或占位凭据、错误的云端点、地址覆盖、端口格式和单/双应用凭据不匹配等常见问题。它不打印配置值、不访问网络，也不修改 Entra。**检查通过不代表管理员同意、Conditional Access 或真实 OAuth/Graph 调用已经验证。**

[常见问题与检查边界](docs/QUICKSTART.md#遇到问题时)

## 客户端兼容性

仓库此前记录 WorkBuddy、Qoder Work、Codex 和 Dify 已完成远程连接、OAuth、工具发现与调用验证；这不是本次维护对这些客户端最新版的重新认证。详细记录保留在[完整部署说明](DEPLOYMENT.md#已验证客户端)。

其他客户端需要支持 Streamable HTTP、OAuth 资源元数据/授权服务器发现、浏览器回调和 Bearer token。客户端能打开登录页面，也仍可能被租户 MFA、设备合规或用户分配策略拦截。

## 安全边界

远程服务使用 HTTPS；生产凭据放入 Secret Manager；只授予实际需要的 Graph delegated permissions 和 App Roles。不要为解决登录问题关闭认证或直接开放全部模块。

使用中国区 Graph **不自动保证端到端数据不出境**：工具结果还会发送给你选择的 AI 客户端/模型服务。上线前需审核客户端数据流、日志留存、模型供应商及组织合规要求。详见[威胁模型](docs/THREAT_MODEL.md)和[安全策略](SECURITY.md)。

## 文档与参与

| 目标 | 入口 |
|---|---|
| 首次跑通、检查配置、逐步启用功能 | [快速上手](docs/QUICKSTART.md) |
| Docker、双应用、角色、生产部署 | [完整部署](DEPLOYMENT.md) |
| 实际使用方式与提示词 | [用户使用说明](USER_GUIDE.md) |
| 工具及所需权限 | [工具目录](docs/TOOL_CATALOG.md) |
| 回调、设备策略与登录问题 | [Entra 登录排障](docs/ENTRA_CONFIDENTIAL_WEB_LOGIN.md) |
| 后续维护优先级与发布记录 | [维护路线](docs/MAINTENANCE.md) · [Changelog](CHANGELOG.md) |

项目解决了你的 21V 接入问题，欢迎点一个 **Star**，也欢迎提交脱敏的使用反馈或客户端兼容性记录。贡献前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)；安全漏洞不要公开提交 Issue。

## License

[Apache License 2.0](LICENSE)。Microsoft、Microsoft 365、Microsoft Entra、Microsoft Graph 等名称属于各自商标权利人，参见 [TRADEMARKS.md](TRADEMARKS.md)。
