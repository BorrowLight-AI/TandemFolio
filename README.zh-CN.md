# TandemFolio

> 一个让你与 AI Agent 在同一份文档上协作的本地优先可视化工作区。

[English](README.md) · [文档导航](docs/README.md) · [快速开始](docs/getting-started.md)

TandemFolio 为 Codex 等 MCP Apps 宿主提供持续存在的可视化编辑能力。你可以先在格式原生的编辑器中打开文档，再让 Agent 操作与你眼前完全相同的文件、选区、撤销历史和版本。

它适合希望获得 Agent 协助、但不愿把文档交给另一套黑盒文件生成流程的人。文件保留在本地，编辑过程可见，Agent 的每次修改都通过已挂载的编辑器完成。

## 演示

这段预览展示了 Codex 打开 TandemFolio 的 PPTX 编辑器，并在同一任务中生成一套包含 10 页、可继续编辑的演示文稿。点击预览可观看完整的 49 秒录像。

[![TandemFolio PPTX 编辑演示](docs/assets/tandemfolio-demo.gif)](docs/assets/tandemfolio-demo.mp4)

[观看或下载完整 MP4 演示](docs/assets/tandemfolio-demo.mp4)

## 它解决什么问题

- **同一份文档，两个协作者。** 人与 Agent 操作同一个实时编辑器状态，不会生成互相脱节的副本。
- **本地、可视化地编辑。** 在 MCP Apps 宿主内直接处理本地 DOCX、XLSX、PPTX、PDF 和 Markdown 文件。
- **可预期的 Agent 修改。** 所有操作都带有明确类型和版本约束，并复用编辑器自身的状态和撤销路径。
- **不依赖产品账号或云端文档服务。** 已打包产品由浏览器编辑器和本地 MCP 服务组成。

## 支持格式

| 格式 | 可视化编辑器 | 常见工作 |
| --- | --- | --- |
| DOCX | 文档画布 | 文本、版式、表格、批注、图片 |
| XLSX | 电子表格画布 | 单元格、公式、工作表、图表 |
| PPTX | 幻灯片画布 | 幻灯片、文本、布局、对象 |
| PDF | PDF 工作区 | 文本、批注、表单、页面 |
| Markdown | 富文本 Markdown 编辑器 | 写作、结构、Frontmatter、导出 |

## 快速开始

Windows/macOS 用户可前往 [GitHub Releases](https://github.com/BorrowLight-AI/TandemFolio/releases)
下载 `tandemfolio-X.Y.Z.zip`（两平台通用）或 `.tar.gz`（macOS），核对 `SHA256SUMS` 后，
按照[安装与更新教程](docs/distribution.md)运行包内的 `install.ps1` 或 `install.sh`。
需要 Node.js 22.12+ 和支持插件的 Codex CLI，无需编译源码或安装项目依赖。
只有当前源码发布证据和 Windows/macOS 冒烟验证通过后，流水线才会提供发布附件。

在源码目录中，用一条命令安装依赖并启动本地 MCP 服务：

```bash
npm install --ignore-scripts && npm run dev
```

需要 Node.js 22.12+ 和 npm 10+。要构建可本地安装的自包含插件，请运行 `npm run build`；宿主安装和独立编辑器运行方式见[快速开始](docs/getting-started.md)。

## 一次协作如何发生

1. 在 MCP Apps 宿主中打开 TandemFolio 编辑器。
2. 宿主为当前文档保持一个实时编辑会话。
3. 请求 Agent 查看或修改文档。
4. Agent 的类型化操作进入同一个已挂载编辑器，更新可见文档及其撤销历史。

编辑器关闭时，修改会失败，不会悄悄改写某个隐藏副本。

## 项目状态

TandemFolio 仍处于预发布阶段。五种格式的编辑器均已打包，但只有当前源码同时通过完整的视觉、性能、冒烟、许可证和仓库证据门禁后，项目才会进入可发布状态。准确状态与来源信息请见[项目事实与归属](docs/project-facts.md)。

## 文档

- [快速开始](docs/getting-started.md) — 从源码运行、构建插件、在宿主中打开。
- [Windows/macOS 安装](docs/distribution.md) — 下载、校验、安装、更新和回退。
- [项目事实与归属](docs/project-facts.md) — 来源、Apache-2.0 义务与修改记录。
- [实时会话协议](docs/protocol/live-session.md) — 已实现的会话、工具、版本与持久化行为。
- [开发指南](docs/development.md) — 构建、测试、打包、发布和排障。
- [架构决策](docs/adr/) — 已接受的技术决策。

## 参与贡献

欢迎参与。请先阅读 [CONTRIBUTING.md](CONTRIBUTING.md)，涉及上游衍生代码时再阅读[项目事实与归属](docs/project-facts.md)。

## 许可证与归属

TandemFolio 以 [Apache License 2.0](LICENSE) 发布，其中包含并修改了 Apache-2.0 社区代码。原始版权声明、许可证文本和项目 [NOTICE](NOTICE) 都被保留；完整来源与修改清单见[项目事实与归属](docs/project-facts.md)。
