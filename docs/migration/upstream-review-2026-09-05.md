# 2026-09-05–06 GenOffice 上游选择性同步

## 范围与结果

- 查询分支：`genspark-ai/genoffice` 的 `main`；fetch 后固定候选为 `360ce0625eaf748368e5535984b073f6fb2487b5`（2026-09-04）。
- 提取基线仍为 `dc4d7e5927864498913b7ba42d0da06cc7cf628e`。两者相差 41 个提交、1,487 个社区变更路径；未检查或引入 `ee/`。
- 按用户确认，完整移植候选中适用于浏览器挂载权威边界的 DOCX、XLSX 与 PPTX 原生编辑、排版、渲染、导入导出与保真能力。采用选择性来源移植，不创建 merge parent，也不把五格式整体源码固定点提升到候选提交。
- 明确排除 AI 面板/工具/提示词、账号、遥测、Electron main/preload、IPC、桌面系统打印、整文件密码加密和企业版来源。现有 DOCX OOXML 文档保护及其 typed MCP 路由保留。
- DOCX 新增 `docx.image.set_z_order`，XLSX 新增计算、保护、分页符、名称和工作簿合并路由，PPTX 新增背景、对象效果/几何和文字框属性路由；生成目录现为 103 个 DOCX 操作、123 个 XLSX 操作、81 个 PPTX 操作、354 个总操作。它们与各自原生 UI 共用挂载编辑器状态、Undo、恢复及保存路径。

## 已合入的 DOCX 原生特性

| 能力组                   | 具体特性与修复                                                                                                                                  | 本地接入方式                                                                                                                              |
| ------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- |
| OOXML 解析与保存         | XML 实体只解码一次；深层嵌套表格；多 body 扫描；AlternateContent/MCE；主部件非标准路径；链接图片；资源清理；ZIP 本地名称；未知部件保留          | 更新 `packages/docx-engine` 的分层 parse/package/normalize/save 管线，并保留 TandemFolio 的图像裁剪、变换、环绕、页眉页脚变体和标题页扩展 |
| 字体与文字保真           | 字体表、主题/双槽字体、CJK/韩文/泰米尔/PUA 字体度量、RTL run、隐藏文字、小型大写/全大写、字符间距、字符单位缩进及保存取消标记                   | 浏览器字体资产随 DOCX renderer 打包；解析、PM 状态和 OOXML save/reopen 回归闭环                                                           |
| 段落与版式               | CJK 标点压缩、Word 2013+ 两端对齐压缩、自动连字符语言、默认制表位、段落边框合并、自动段间距清除、样式应用、大小写转换、行距/缩进/悬挂缩进快捷键 | 文件打开时注入文档兼容设置；Ribbon、对话框和快捷键调用同一原生 transaction                                                                |
| 分页与分节               | 显式/连续/奇偶/下一栏分节；混合栏宽和 RTL 栏序；竖直居中/底部对齐；不同节纸张宽度/页边距；字符/行网格；删除态分节符；空白页；页码格式           | 主画布、实时分页、TOC 页码和分页预览共享 `liveSections`、切片和每块布局装饰；#177 的百分比表格及分节边距已接通                            |
| 跨页表格                 | 重复表头、不可拆行、单元格内分页、声明行高补齐、超宽表格页边缘约束、网格协调、右键插入/删除/选择/合并/拆分和表格属性                            | 使用 PM tables 原生命令与 renderer 自有表格状态；现有 typed 表格操作和 Undo 不变                                                          |
| 图片、浮动对象与形状     | 浮动对象跨页落位、页/边距相对锚点、上下环绕带、前后层级、旋转/翻转、文本框/WordArt/VML、形状文本格式、SmartArt/OLE/EMF/WMF 显示                 | 新层级命令接受精确 block 与 32 位 rank；右键 Arrange 和 MCP 共用 helper；引擎保留原始 DrawingML/VML                                       |
| 页眉页脚、水印与页面边框 | 奇偶页/首页变体；每节引用；富文本、表格、图片和浮动水印；PAGE/NUMPAGES；页边框；不同纸宽的页眉页脚对齐                                          | 首页面浮动图使用独立零高 widget，其余页挂入 page-gap；页面级宽度和位置按所属节计算                                                        |
| 脚注、尾注与审阅显示     | 页底脚注区、文末尾注区、富段落和分隔线；跨段评论锚点；边距批注；非活动选择；修订视图中的删除分节符                                              | 注释/notes 与分页共用稳定 ID 和挂载文档状态；原有 typed 评论/修订生命周期保留                                                             |
| 原生编辑体验             | 查找替换焦点、Word Count、快捷键帮助、Word 风格快捷键、单单元格电子表格粘贴解包、图片右键命中、共享 Dropdown、符号字体显示、可访问 tooltip      | 移除 F7 AI 校对入口；所有保留动作都落到 renderer 自有状态，纯 UI 导航不新增 MCP 变更命令                                                  |

同时保留此前已移植的 PPTX 旋转/翻转表格命中与串行保存、XLSX 大范围分批读取修复。它们不是本次 DOCX 扩展的依赖。

## 已合入的 PPTX 原生特性

| 能力组              | 具体特性与修复                                                                                                                                 |
| ------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| OOXML 与保存保真    | 主题覆盖和默认文字样式继承、显式关闭字符属性、段落制表位、符号字体、背景继承、隐藏形状、占位图片几何、组填充、关系与媒体资源清理、未知部件保留 |
| 图表与 SmartArt     | ChartEx 解析/回退、稀疏缓存、坐标轴/图例/数据标签与线型保真、`bar3D`/`pie3D` 插入更新、Diagram 层级与 SmartArt 回退/保存                       |
| 形状、WordArt 与 3D | 预设/自定义几何、调整手柄与形状转换、图案/渐变/图片填充、渐变线和线帽/连接/复合线、阴影/发光/倒影/柔化边缘、Scene 3D/倒角、WordArt 效果        |
| 文字与 RTL          | 横排、竖排、270 度、堆叠和 WordArt 竖排，自动适应/换行/内边距，文字高亮和变形，东亚/韩文字体识别，RTL 段落与表格方向，编辑后显式格式保留       |
| 图片、表格与交互    | 图片裁剪内缩与 blip 效果、tile/stretch 图片填充、旋转/翻转表格单元格命中、RTL 调宽、中心旋转支点、调整拖动合并为一个 Undo 单元                 |
| 字体、背景与输出    | 打开时在首帧前注册 OOXML 嵌入字体；纯色/渐变/图片/重置背景与隐藏母版图形；幻灯片/讲义/备注打印布局；滚轮翻页和当前页放映快捷键                 |
| 原生 UI             | 接入新版格式窗格、背景窗格、形状/图表图库、Office 颜色与最近颜色，以及 19 个语言的应用/Ribbon/窗格拆分字典                                     |

新增七个 typed MCP 操作；既有填充、线条、段落、表格样式和图表操作扩展到完整渐变、效果、RTL 与 3D 字段。真实 BrowserPresentation 集成测试覆盖原生 Undo、单调 revision、保存和重开。

## 未移植项及原因

| 上游区域                                                                 | 处置                                                                                                        |
| ------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------- |
| `renderer/ai`、AiPanel/AiAsk、Agent tools、AI 翻译/校对/改写、模型和搜索 | 用户明确排除，且违反产品无模型流边界                                                                        |
| Electron main/preload、IPC、shell 生命周期、系统字体/系统打印宿主        | 浏览器/MCP host 已是唯一宿主；引入会产生第二条文件或文档权威路径                                            |
| 整文件密码加密打开/保存、桌面 PrintDialog                                | 上游实现依赖 Electron/IPC；不进入浏览器 renderer。OOXML 内部文档保护仍保留并可编辑                          |
| `ee/`                                                                    | 禁止检查和导入                                                                                              |
| 上游 i18n 目录拆分、共享 font-list 的结构性搬家                          | 本地保留等价的 19 语言表和浏览器字体枚举；Office ColorPicker 已按浏览器宿主边界接入 XLSX，不引入 AI catalog |

DOCX、XLSX 与 PPTX 候选范围内没有剩余待移植的适用原生能力。XLSX 已完成全文件查找、筛选一致性、错误扫描、活动行列高亮、CSV 导出、工作簿合并、RTL、OLE、Office 配色、1904 日期系统、手动分页符、分页预览、精确打印设置和首页/奇偶页页眉页脚布局。后续清单只保留 PDF 评论/文字编辑保真/阅读位置、Markdown 数学公式与资源生命周期。

## 逐提交处置

| 提交                                                               | 日期       | 上游标题                                                                                     | 本次处置                                                                                                           |
| ------------------------------------------------------------------ | ---------- | -------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| [945c370](https://github.com/genspark-ai/genoffice/commit/945c370) | 2026-08-13 | fix(docx): decode XML entities exactly once in parsed run text (#89)                         | 已移植：DOCX XML 实体只解码一次；含编辑后保存回归。                                                                |
| [d5558b6](https://github.com/genspark-ai/genoffice/commit/d5558b6) | 2026-08-13 | Sync snapshot (2026-08-13) (#92)                                                             | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [e4d9d40](https://github.com/genspark-ai/genoffice/commit/e4d9d40) | 2026-08-16 | Replace unrecognizable SaveIcon with floppy disk (#96)                                       | 本地已使用统一保存图标；不覆盖本地窄屏工具栏。                                                                     |
| [04a994b](https://github.com/genspark-ai/genoffice/commit/04a994b) | 2026-08-17 | Sync snapshot (2026-08-16) (#99)                                                             | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [fe2e174](https://github.com/genspark-ai/genoffice/commit/fe2e174) | 2026-08-19 | Sync snapshot (2026-08-19) (#110)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [f68df70](https://github.com/genspark-ai/genoffice/commit/f68df70) | 2026-08-21 | Sync snapshot (2026-08-20) (#123)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [6d9b681](https://github.com/genspark-ai/genoffice/commit/6d9b681) | 2026-08-23 | Sync snapshot (2026-08-23) (#130)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [afc6711](https://github.com/genspark-ai/genoffice/commit/afc6711) | 2026-08-24 | fix(sheets): extend Ctrl+F beyond the loaded window on streamed workbooks (#131)             | 已移植：浏览器分批扫描未加载行，和 Univer 查找/替换、筛选及原生 Undo 汇合。                                        |
| [cc8cff4](https://github.com/genspark-ai/genoffice/commit/cc8cff4) | 2026-08-24 | feat(sheets): cross-highlight the active cell's row and column (#132)                        | 已移植：活动行列高亮使用一个有界 Univer Canvas 扩展及本地视图偏好。                                                |
| [5073e4d](https://github.com/genspark-ai/genoffice/commit/5073e4d) | 2026-08-24 | Sync snapshot (2026-08-24) (#143)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [da3b1ca](https://github.com/genspark-ai/genoffice/commit/da3b1ca) | 2026-08-25 | fix(slides): sweep generated-page temp files with a TTL at startup (#141)                    | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [bc1dceb](https://github.com/genspark-ai/genoffice/commit/bc1dceb) | 2026-08-25 | fix(sheets): row-batch over-cap sidecar reads so streaming survives big viewports (#140)     | 已移植：XLSX 大视口读取按 18,000 单元格预算分批，索引落后时停止。                                                  |
| [7a814db](https://github.com/genspark-ai/genoffice/commit/7a814db) | 2026-08-25 | fix(sheets): keep filter-hidden rows out of full-sheet find results (#133)                   | 已移植：全文件查找合并文件命中时排除筛选隐藏行。                                                                   |
| [f81dd3d](https://github.com/genspark-ai/genoffice/commit/f81dd3d) | 2026-08-26 | fix(markdown): wait for in-flight saves during close instead of silently failing (#149)      | 不直接移植：修复 Electron 关闭回调；本地已无该关闭协议。                                                           |
| [b1a01e2](https://github.com/genspark-ai/genoffice/commit/b1a01e2) | 2026-08-26 | fix(ai-provider): release SSE reader when the consumer abandons the stream mid-flight (#150) | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [72a4a42](https://github.com/genspark-ai/genoffice/commit/72a4a42) | 2026-08-26 | fix(ai-provider): skip malformed SSE frames instead of killing the AI turn (#151)            | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [9792915](https://github.com/genspark-ai/genoffice/commit/9792915) | 2026-08-26 | feat(slides): flag horizontal text overflow in the layout audit (#154)                       | 排除 AI 布局审计入口；原生文字布局修复另行评估。                                                                   |
| [2e3e97f](https://github.com/genspark-ai/genoffice/commit/2e3e97f) | 2026-08-26 | feat(slides): add a set_speaker_notes agent tool (#153)                                      | 不导入 AI tool；本地已有 pptx.notes.set 与原生备注编辑。                                                           |
| [7eb5d59](https://github.com/genspark-ai/genoffice/commit/7eb5d59) | 2026-08-26 | fix(sheets): scan the whole file for error checking on streamed workbooks (#134)             | 已移植：错误检查使用浏览器文件分页扫描、结构映射和挂载选择导航。                                                   |
| [9711a45](https://github.com/genspark-ai/genoffice/commit/9711a45) | 2026-08-26 | fix(shell): surface project IPC failures and AI settings save/test errors (#156)             | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [0a2c25d](https://github.com/genspark-ai/genoffice/commit/0a2c25d) | 2026-08-26 | Sync snapshot (2026-08-26) (#159)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [607c770](https://github.com/genspark-ai/genoffice/commit/607c770) | 2026-08-27 | feat(shell): open documents dropped onto the window (#161)                                   | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [583a045](https://github.com/genspark-ai/genoffice/commit/583a045) | 2026-08-27 | feat(markdown): prompt to save before pasting images into untitled docs (#162)               | 不适用：本地粘贴图片内嵌 data URL，不要求未命名文档先落盘。                                                        |
| [2239cce](https://github.com/genspark-ai/genoffice/commit/2239cce) | 2026-08-30 | Sync snapshot (2026-08-30) (#168)                                                            | DOCX 非 AI 原生部分已选择性移植；其他格式和宿主部分仍按边界排除或后续评估。                                        |
| [3548a80](https://github.com/genspark-ai/genoffice/commit/3548a80) | 2026-08-31 | fix: rotated table cell dblclick-to-edit in slides (#171)                                    | 已适配：旋转/翻转表格命中与编辑框；复用已有单元格编辑/MCP 路由。                                                   |
| [566c1f3](https://github.com/genspark-ai/genoffice/commit/566c1f3) | 2026-09-02 | fix(agent-core): discard stale compaction results after reset (#178)                         | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [93b8938](https://github.com/genspark-ai/genoffice/commit/93b8938) | 2026-09-02 | fix docs per-section page margins and pct table widths (#177)                                | 已移植：`sectionWidthSpecs`、块位移、每节边距变量、百分比表格约束及 gap 页眉页脚对齐已接入实时分页。               |
| [9f971ed](https://github.com/genspark-ai/genoffice/commit/9f971ed) | 2026-09-03 | Sync snapshot (2026-09-02) (#181)                                                            | DOCX、XLSX 的浏览器安全非 AI 原生能力已选择性移植；Electron/IPC、原生侧车、AI 与其他格式部分按边界排除或后续评估。 |
| [99d376b](https://github.com/genspark-ai/genoffice/commit/99d376b) | 2026-09-03 | fix markdown/pdf ai panels surfacing a truncated reply as "no reply" (#184)                  | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [f105f36](https://github.com/genspark-ai/genoffice/commit/f105f36) | 2026-09-03 | feat ai settings for the per-turn output token cap (#185)                                    | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [e833fff](https://github.com/genspark-ai/genoffice/commit/e833fff) | 2026-09-03 | fix(slides): skip width overflow check on vertical text bodies (#160)                        | 排除 AI 布局审计入口的竖排误报修复；不新增上游 AI 审计。                                                           |
| [2dab773](https://github.com/genspark-ai/genoffice/commit/2dab773) | 2026-09-04 | fix(ai-provider): sanitize Gemini tool schemas for union types and refs (fixes #186) (#187)  | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [6f71780](https://github.com/genspark-ai/genoffice/commit/6f71780) | 2026-09-04 | fix(sheets,slides): forward stopReason so max_tokens truncation recovers (#190)              | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [f30a301](https://github.com/genspark-ai/genoffice/commit/f30a301) | 2026-09-04 | fix(ai-provider): return ok:false on non-JSON 200 instead of throwing (#191)                 | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [eec2737](https://github.com/genspark-ai/genoffice/commit/eec2737) | 2026-09-04 | fix(ai-panels): clear preview marker on read rejection so thumbnails retry (#192)            | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [4846ace](https://github.com/genspark-ai/genoffice/commit/4846ace) | 2026-09-04 | fix(slides): serialize concurrent saves to prevent file corruption (#155)                    | 已适配：保存串行化；扩展覆盖本地 Save As，失败后队列可继续。                                                       |
| [4770d9c](https://github.com/genspark-ai/genoffice/commit/4770d9c) | 2026-09-04 | fix(ai-provider): keep default temperature for the Gemini 3 family (#193)                    | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [2910486](https://github.com/genspark-ai/genoffice/commit/2910486) | 2026-09-04 | feat(ai-search): add Tavily web search fallback via TAVILY_API_KEY (#194)                    | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [d8e233c](https://github.com/genspark-ai/genoffice/commit/d8e233c) | 2026-09-04 | docs(contributing): document TAVILY_API_KEY in env var table (#207)                          | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [9439d3e](https://github.com/genspark-ai/genoffice/commit/9439d3e) | 2026-09-04 | feat(docs): RTL-aware AI panel chrome and auto-direction messages (part of #13) (#202)       | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |
| [360ce06](https://github.com/genspark-ai/genoffice/commit/360ce06) | 2026-09-04 | fix(slides): preserve agent runs across lifecycle boundaries (#208)                          | 排除：AI、桌面壳、生命周期/设置或其文档，不属于本次原生编排范围。                                                  |

## 验证与发布状态

- PPTX 局部验证：engine 81 个文件/836 项、render 10 个文件/245 项、Slides 21 个文件/266 项通过；三个工作区类型检查通过。
- 生成 Manifest 为 354 个操作，其中 DOCX 103 个、XLSX 123 个、PPTX 81 个；确定性 Manifest、全仓类型检查、构建和发布门禁状态以本次最终验证记录为准。
- 各根工作区测试通过：377 个测试文件、4,714 项；另有 1 个环境条件文件/断言跳过。XLSX 的 174 个文件、1,984 项和 MCP server 的 13 个文件、386 项全量通过。
- `npm test -w @genoffice/docx-engine`：87 个文件通过，996 项通过、1 项跳过；覆盖候选的 OOXML 解析、保存、复杂排版和对象保真。
- `npm test -w @genoffice/docs`：147 个文件、1,680 项通过；包括从 AI tool 解耦后的字符单位缩进 UI→保存→重开用例。
- `npm run build`：五格式与 MCP 插件重新构建通过。生成资源保留在本地工作区供开发验证，并非获准发布的归档。
- `npm run smoke:mcp`：26 个总工具（11 个公开、15 个 app-only）冒烟通过，包括 `office_merge_local_workbook`。
- `npm run measure:assets`：五格式通过。DOCX 实测 3,508,609 raw / 987,899 gzip，XLSX 实测 20,867,055 / 8,532,458，均低于各自 raw 上限。`npm run licenses` 与 `npm run upstream:check` 通过。
- `npx playwright test --config playwright.host.config.ts tests/visual/docx-community.spec.ts`：5 项真实浏览器测试通过，覆盖可访问原生工具栏、File 菜单、Session 绑定保存、离屏释放/恢复和恢复版本单调性。
- `npm run release:gate` 按设计以 `source_mismatch` 阻断，写入 `ready: false`；未改证据哈希、阈值或批准状态绕过发布检查。

原生 renderer/engine 源码变化后，旧批准证据不再 source-current；生成的 `release-readiness.json` 已由门禁置为 `ready: false`。正式发布仍需重新采集五格式性能/视觉证据，不把功能回归通过当作正式发布批准。本次未修改证据、阈值、掩码或批准状态来规避检查。
