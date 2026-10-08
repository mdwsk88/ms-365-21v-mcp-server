# 使用已发布的 Docker 镜像

[首页](../README.md) · [版本下载](https://github.com/mdwsk88/ms-365-21v-mcp-server/releases) · [完整部署](../DEPLOYMENT.md)

镜像：`ghcr.io/mdwsk88/ms-365-21v-mcp-server:v0.1.0`，支持 `linux/amd64` 和 `linux/arm64`。Docker 自动选择对应架构，无需在本机编译服务。建议固定版本标签；`latest` 会随版本发布更新。

## 首次启动

1. 按[快速上手](QUICKSTART.md)配置 Entra 应用、Web 回调、Graph delegated 权限和用户 App Roles。
2. 获取版本文件，运行配置向导；向导需要 Node.js 22+，无需先安装 npm 依赖：

   ```bash
   git clone --branch v0.1.0 https://github.com/mdwsk88/ms-365-21v-mcp-server.git
   cd ms-365-21v-mcp-server
   npm run setup
   ```

3. 编辑生成的 `.env`，填写 `MS_CLIENT_SECRET` 的 Value。保留已有配置时直接使用原文件。检查后启动：

   ```bash
   npm run doctor
   docker compose -f docker-compose.release.yml pull
   docker compose -f docker-compose.release.yml up -d
   docker compose -f docker-compose.release.yml ps
   docker compose -f docker-compose.release.yml logs --tail 50
   ```

没有 Node.js/Git 时，也可从 Release 下载 `docker-compose.release.yml` 和 `env.example`，将后者命名为 `.env`，按[完整部署](../DEPLOYMENT.md)手动填写配置，再执行上面的 Docker 命令。完整示例可能启用更多模块，首次试用建议设置 `MCP_TOOL_CATEGORIES=users`，只开放个人资料查询。

Compose 将容器以非 root 用户运行，并使用命名卷保存 `/app/.tokens`，避免新建 Linux 宿主目录的写权限问题。只发布宿主机 `127.0.0.1:3000`，本机桌面客户端连接 `http://localhost:3000/mcp`。修改端口时设置 `.env` 的 `MCP_DOCKER_PORT`，同时更新公网地址和 Entra 回调。

远程客户端应通过宿主机 HTTPS 反向代理连接。设置 `MCP_PUBLIC_BASE_URL=https://你的域名`，Entra Web 回调为 `https://你的域名/oauth/microsoft/callback`。代理需转发 MCP、OAuth 和 `/.well-known/` 路由。代理若也运行在容器中，应将其与 MCP 加入共享网络并转发到 `mcp-gateway:3000`。具体配置见[完整部署](../DEPLOYMENT.md)。

连接成功后测试 `auth_status` 和 `graph_get_me`；`healthy` 只代表服务已启动。

## 更新与回退

发布说明会列出对应代码、镜像版本和验证情况。升级前保留旧镜像标签、备份 `.env` 与状态卷；在 `.env` 设置 `MCP_IMAGE_TAG=目标版本`，然后重新执行 `pull` 和 `up -d`。回退时改回旧标签并执行 `up -d`。需要严格固定镜像时，可以把 Compose 的 `image` 改为 Release `image-digest.txt` 中的完整引用。

正常 `docker compose down` 保留命名卷；`down -v` 会删除登录状态，请勿用于常规升级。由原源码/绑定目录部署迁移过来时，应先停旧服务、备份并迁移其 `.tokens`，检查容器用户的读写权限。直接使用新卷会要求用户重新登录。

镜像只打包程序及运行依赖；凭据通过 `.env` 注入，不应烘焙进镜像。此次镜像发布不会修改既有 AWS、AgentRun 或 Entra 配置。
