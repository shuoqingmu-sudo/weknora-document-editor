# Weknora 文档编辑器

`weknora-document-editor` 是一个面向 Weknora 智能体的远程 MCP 文档服务。它让智能体能够将对话中的内容整理为 Word、Excel 或 PDF 文件，并把文件作为可下载附件交付给用户。

服务采用 MCP SSE 传输协议，适合部署到 ModelScope 或任何支持 Docker 的平台。与本地工具不同，远程服务不会返回宿主无法访问的服务器文件路径，而是返回带随机令牌的 HTTPS 下载链接。Weknora 获取该链接后即可将真实文件附加到当前对话。

## 能力

| MCP 工具 | 输出格式 | 适用场景 |
| --- | --- | --- |
| `create_word` | `.docx` | 报告、方案、会议纪要、通知 |
| `create_excel` | `.xlsx` | 台账、数据清单、统计表、多工作表导出 |
| `create_pdf` | `.pdf` | 面向客户的摘要、归档版本、可打印报告 |

三个工具均支持标题、段落和表格；Excel 工具支持多工作表。PDF 使用 Unicode 字体，可生成中文内容。

## 工作方式

```text
Weknora 对话 -> MCP SSE 调用 -> 文档生成服务 -> 令牌化下载 URL -> Weknora 附件 -> 用户下载
```

每次生成都使用随机令牌命名文件。工具响应包含：

- `resource_link.uri`：可由 Weknora 下载并转为对话附件的 HTTPS URL。
- `structuredContent.url`：同一文件的机器可读 URL。
- `filename` 和 `mimeType`：用于正确显示和上传附件。

## MCP 端点

| 端点 | 方法 | 说明 |
| --- | --- | --- |
| `/health` | `GET` | 健康检查，返回服务名称与状态 |
| `/sse` | `GET` | MCP SSE 入口 |
| `/messages?sessionId=...` | `POST` | SSE 会话的 MCP JSON-RPC 消息入口 |
| `/files/<token>/<filename>` | `GET` | 生成文件的下载地址 |

`/sse` 和 `/messages` 在配置 `MCP_API_KEY` 后需要 Bearer Token。文件下载 URL 自身含随机令牌，便于 Weknora 的附件上传器直接获取文件。

## 本地运行

需要 Docker。先构建镜像：

```bash
docker build -t weknora-document-editor .
```

再启动服务：

```bash
docker run --rm -p 7860:7860 \
  -e PUBLIC_BASE_URL=http://localhost:7860 \
  -e MCP_API_KEY=replace-with-a-long-random-secret \
  -v weknora-documents:/data \
  weknora-document-editor
```

验证服务：

```bash
curl http://localhost:7860/health
```

预期响应：

```json
{
  "status": "ok",
  "server": "weknora-document-editor-sse"
}
```

## Weknora 接入

在 Weknora 的远程 MCP 配置中使用 SSE 传输。将域名和密钥替换为部署后的真实值：

```json
{
  "transport": "sse",
  "url": "https://YOUR-MODELSCOPE-DOMAIN/sse",
  "headers": {
    "Authorization": "Bearer YOUR_MCP_API_KEY"
  }
}
```

Weknora 的附件工作流应读取 MCP 响应中的 `resource_link.uri` 或 `structuredContent.url`，下载文件后作为当前会话附件发送。不要将 URL 仅作为普通文本回复给用户。

## 工具调用示例

下面是创建 PDF 周报的 MCP 参数示例：

```json
{
  "name": "create_pdf",
  "arguments": {
    "filename": "weekly-status",
    "title": "项目周报",
    "paragraphs": [
      "本周完成 API 联调，发布计划保持不变。"
    ],
    "tables": [
      {
        "headers": ["模块", "状态", "负责人"],
        "rows": [
          ["接口", "已完成", "张三"],
          ["前端", "进行中", "李四"]
        ]
      }
    ]
  }
}
```

创建 Excel 时，使用 `sheets` 数组替代 `title`、`paragraphs` 和 `tables`：

```json
{
  "name": "create_excel",
  "arguments": {
    "filename": "project-data",
    "sheets": [
      {
        "name": "进度",
        "headers": ["任务", "状态"],
        "rows": [["需求评审", "完成"], ["联调", "进行中"]]
      }
    ]
  }
}
```

## 部署到 ModelScope

1. 在 GitHub 创建仓库，并将本目录作为仓库根目录推送。
2. 在 ModelScope 创建 Docker 类型的部署或 Space，并关联 GitHub 仓库。
3. 将服务端口设置为 `7860`。
4. 配置环境变量：

| 变量 | 必填 | 说明 |
| --- | --- | --- |
| `PUBLIC_BASE_URL` | 是 | ModelScope 分配的完整公网 HTTPS 域名，例如 `https://example.modelscope.cn` |
| `MCP_API_KEY` | 强烈建议 | 用于保护 MCP SSE 入口的长随机密钥 |
| `PORT` | 否 | 服务端口，默认 `7860` |
| `OUTPUT_DIR` | 否 | 文件输出目录，默认 `/data/generated` |

5. 可选地将持久化存储挂载到 `/data`，避免重启后丢失已生成的文件。
6. 部署完成后访问 `https://YOUR-MODELSCOPE-DOMAIN/health`。健康检查成功后，将 `/sse` URL 填入 Weknora。

## 安全与运维

- 公网部署必须设置 `MCP_API_KEY`，并只在 Weknora 的 MCP 密钥配置中保存该值。
- `PUBLIC_BASE_URL` 必须是用户和 Weknora 都能访问的 HTTPS 地址；它决定响应中下载链接的主机名。
- 生成文件包含随机 URL 令牌，但当前版本不会自动过期或清理文件。生产环境建议为 `/data/generated` 配置定期清理策略和容量监控。
- 文件内容由上游智能体输入。请在 Weknora 侧设置适当的内容审核、身份认证和访问权限。

## 项目结构

```text
.
├── Dockerfile       # 容器构建与运行入口
├── requirements.txt # 文档生成依赖
├── server_sse.py    # MCP SSE 服务与文件下载端点
└── README.md        # 部署与接入说明
```

## 开发说明

项目使用 Python 标准库实现 SSE 与 MCP JSON-RPC，没有绑定特定云平台 SDK，因此可在 ModelScope、任意 Docker 平台或内网服务器运行。运行依赖仅包括 `python-docx`、`openpyxl` 和 `reportlab`。
