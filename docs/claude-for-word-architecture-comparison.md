# Claude for Word 与 Claric：架构比较与改进建议

分析日期：2026-10-30。Claric 基线：`6f1c41c`，包版本 `1.0.2`。

后续补充：[本机缓存静态分析](claude-cross-document-local-analysis.md)已确认该缓存构建的 conductor 协议、会话增量共享、虚拟文件系统和目标本地 Agent 委派。本文原有“内部未知”的范围应结合该报告缩小；服务端实现和实际运行行为仍未验证。

本文依据 Claude 官方文档、Microsoft Marketplace、Anthropic 公开 manifest 和 Claric 当前源码。Claude 的内部执行器、文档表示、提示词和 diff 实现没有公开；涉及这些内部机制的描述只作为设计推断。本文未在本机已登录的 Claude 插件中发起文档编辑，也未进行速度、质量或格式保持的对照测试。官方当前能力不等于本机缓存版本已经具备全部能力。

## 核心判断

两者共享 Office Web Add-in 的基本形态。Claude 的主要可见优势是文档问答、引用导航、批注处理、模板编辑和跨 M365 工作流的产品整合；Claric 的主要架构资产是开放模型接入、可自行托管、明确的 proposal 生命周期和可检查的 Word 写入保护。

Claric 已经具备文档工具代理、任务依赖图、资源锁、跨 Word 文档协调、读写验证和视觉读取入口。建议在这些基础上统一文档身份、定位、草稿和上下文，而不是重做一个通用 Agent 或另起任务运行时。

## 1. 部署与数据路径

### Claude：托管前端，多种推理连接路径

直接读取官方 [manifest-word.xml](https://pivot.claude.ai/manifest-word.xml)，得到：

- `ProviderName=Anthropic`，`Host=Document`。
- `SourceLocation=https://pivot.claude.ai/?m=word-1.0.0.1`。
- `Permissions=ReadWriteDocument`，声明 `SharedRuntime 1.1`。

因此，本机安装的是由 manifest 指向托管网页的 Word 插件入口。安装不意味着模型在本机运行；manifest 的 `1.0.0.1` 也不能作为托管前端构建版本或功能集合的可靠标识。[官方存储文档](https://claude.com/docs/office-agents/data-storage)进一步说明 Office 会将 SourceLocation 放入内嵌浏览器运行。

Claude 账号模式将推理请求送往 Anthropic；企业还可使用 Bedrock、Vertex AI、Foundry 或 LLM gateway。网关和云平台连接不等于可以自行托管插件前端，第三方连接模式也有功能差异。[第三方平台文档](https://claude.com/docs/office-agents/third-party-platforms)

### Claric：可托管前端，多模型适配，本地代理可选

`manifest.template.xml` 同样声明 Word taskpane 与 `ReadWriteDocument`。`src/lib/providers.js` 根据部署来源选择直连或同源代理；`scripts/docker-server.cjs` 提供静态资源服务和显式启用的模型代理。模型包含云 API、Ollama、vLLM 等，Anthropic Messages 与 OpenAI-compatible 请求由 `src/lib/llm-client.js` 适配。

这给予 Claric 更大的部署和模型选择空间，同时增加证书、CORS、混合内容、API 差异及模型能力验证的维护成本。静态云端托管前端、云模型调用和 MCP 调用都会影响实际数据边界，不能笼统宣称“Claric 全部本地”或“数据绝不外传”。

## 2. 当前架构对照

| 维度 | Claude 官方可确认的能力 | Claric 当前实现与判断 |
| --- | --- | --- |
| Word 集成 | 托管 taskpane，读写当前文档 | 同类 Office.js 集成，可控制前端部署 |
| 模型接入 | Claude 模型；账号或企业云平台/网关路径 | 多厂商、本地模型和自定义端点；灵活性更大，兼容性负担也更大 |
| 文档问答 | 点击章节引用定位；按主题查找内容 | 通常提取文档上下文后流式回答；已有部分 chunk/file 引用入口，尚未统一为问答级原生证据导航 |
| 多步编辑 | 面向完整文档工作流的公开产品行为；内部编排未知 | 路由、能力目录、planner、任务图、多个专用执行器共同工作 |
| 草稿与写入 | 选区改写、修订模式和模板编辑 | 多种 proposal；默认手动 Apply，可选 auto-apply；基线、锚点、结构指纹和写后检查 |
| 批注闭环 | 读取线程、修改锚定内容并回复 | 能抽取线程、生成/插入批注、分析汇总，也有删除批注专用能力；目录明确不支持通用回复、解决和筛选管理 |
| 复杂结构 | 官方称保持样式、编号及模板表格布局；实际边界需测试 | 表格、图片、格式、页设置有专用能力；通用正文草稿将标题、列表、表格等视为受保护结构 |
| 扩展 | Connectors 与 Skills | 已有 MCP、资源工具、Skills 与本地文件库，不是缺失项 |
| 跨文件 | 多个已打开 M365 应用之间工作；第三方平台模式不支持跨应用 | 已有跨 Word 文档协调；不是零基础，也尚不等于 Excel/PowerPoint/Outlook 适配 |
| 会话 | 本机 IndexedDB 历史，自动压缩；不跨设备同步 | 聊天主要存 localStorage；文件库用 IndexedDB；模型历史主要按预算保留近期回合 |

Claude 能力依据：[Word 指南](https://claude.com/docs/office-agents/word)、[Marketplace](https://marketplace.microsoft.com/en-us/product/wa200010453)、[Connectors 与 Skills](https://claude.com/docs/office-agents/connectors-and-skills)、[跨 M365 应用](https://claude.com/docs/office-agents/work-across-apps)、[存储说明](https://claude.com/docs/office-agents/data-storage)。表格不构成可靠性或性能排名。

## 3. Claric 源码中已经具备的关键基础

1. **受约束的文档代理。** `src/lib/document-edit-session.js` 提供 outline/search/read/contract/stage/read_draft/validate 的循环，要求最新草稿通过审查后才能完成；`src/taskpane/document-edit-actions.js` 负责原生读取、锚定与 Apply。它已经超出单次“全文发模型再替换”的模式。
2. **任务图与资源保护。** `src/lib/task-runtime/task-graph.js` 验证依赖和环，管理任务结果、取消与资源锁；`src/taskpane/conversation.js` 的 `runCompoundTurn` 已接入。当前执行循环仍逐个 await，DAG 主要表达依赖与阻塞，不代表节点已经并行运行。
3. **默认分离生成与应用。** 配置默认 `autoApplyChanges=false`。启用自动应用是已有模式，因此不能再将“所有写入都必须逐卡点击”描述为绝对不变量。
4. **精细写入保护。** `src/lib/word-mutation-verification.js` 检查写后的当前文本与部分原生结构；`document-edit-actions.js` 使用 bookmark 和结构指纹。这是应当保留和集中管理的资产，但不证明所有操作都具备相同保障或所有宿主都能完整记录修订。
5. **跨文档协议。** `src/taskpane/document-agent.js`、`src/lib/coordination/protocol.cjs` 等实现文档身份、上下文请求、提案、租约和结果消息；远端任务能力仍比本地目录窄，文档快照也有摘录上限。
6. **视觉入口。** `src/taskpane/word-render-tools.js` 已支持可用宿主上的 PDF 渲染读取和用户共享 Word 窗口的读取。缺口是把视觉结果纳入模板与格式修改的验收，而非新增一个截图工具。

`docs/turn-flow-architecture-deep-analysis.md` 中部分“当前缺少任务图”的描述已经落后于实现，不能直接沿用为本次诊断。

## 4. 优势、代价与尚未确定的事项

### Claude

可见优势在于任务闭环：用户可以从问题定位到章节，从批注定位到被评论段落，再完成修改和线程回复；模型、连接器、会话和企业配置形成较统一的体验。这是产品整合优势，不能据此推断其内部一定只有一个执行器或一个共享草稿。

主要代价是对托管前端、产品账号/企业配置和服务可用性的依赖；可配置企业推理路径仍不等于离线部署。跨应用需要对应文件和插件已打开，且连接方式会影响可用功能。[跨应用限制](https://claude.com/docs/office-agents/work-across-apps)

本机历史不等于只有本机保存：Claude 账号模式的后台数据保留与本地历史删除是不同机制。企业自定义 OpenTelemetry collector 还可能接收包含文档内容的完整审计 span。应按具体连接模式和日志配置讨论隐私，不能简单以“云端”判定全部数据路径。[Word 数据处理](https://claude.com/docs/office-agents/word)、[企业审计与分析](https://claude.com/docs/office-agents/enterprise-readiness)

### Claric

可见优势是模型和部署可控、变更过程可检查、专用 Word 写入策略可演进，以及适合继续深化的中英文和科研文档能力。Word diff、原生对象保护、表格工具和显式失败状态值得继续投资。

主要代价是能力被多个执行路径分割：正文、格式、表格、图片和批注的状态、上下文、定位和完成条件并不完全一致。增加一个功能往往要同时改路由、目录、planner、runner、卡片和实际 Word 操作。`conversation.js` 与 `word-actions.js` 的体量也反映出协调职责仍较集中。

不能仅凭源码断言 Claric 比 Claude 更准确、更快、更安全或更善于保持格式。这些需要相同文档、指令、模型条件与宿主上的对照测量。

## 5. 最值得改的架构接缝

### A. 统一文档证据与定位，优先改善问答

`word-actions.js:answerQuestion` 通常将整篇提取文本加入请求。它没有复用正文代理的按块检索与读取流程；`document-model.js:search` 是大小写无关的子串匹配，并非 embedding 检索。`revealTextSnippet` 搜索文本后选择首个匹配，重复句子存在歧义。

建议新增公共 DocumentContext 服务，使用稳定目标描述：`documentId + snapshotId + blockId/objectId + version + locator`。回答引用必须指向实际读取过的证据，并通过同一 resolver 跳转；定位失效或多义时返回显式状态。先实现章节索引、关键词/同义词扩展和重排，再按基准结果决定是否引入本地 embedding。语义检索不必预设为云向量数据库。

### B. 保留专用执行器，统一草稿和变更契约

现有能力目录明确提示：不同管线不能看到另一张尚未应用的 proposal。compound 已会阻塞重叠或有依赖的任务，等待 Apply 后继续。这是安全且诚实的行为，但用户为了完成一个连贯任务可能要多次中断并应用。

建议让执行器逐步输出公共 ChangeSet，至少包含目标、基线、读写集、变更、证据、验证与风险。先打通“正文插入 + 新段落格式”，再增加正文与现有格式的共享 overlay。表格与图片仍由各自原生执行器编译写入，避免把 OOXML 细节交给模型。

统一 writer 应在既有 Apply 锁和 freshness 检查上收拢职责；它不是从零增加一把锁。Word 的多个写入不具备数据库级事务，应记录逐操作执行结果，区分已验证、部分完成和不确定，避免宣传“全部原子回滚”。

### C. 将批注从文本附件升级为任务对象

建立 `commentId/threadId → anchor → requestedAction → proposal → verifiedChange → reply` 的状态链。读取不应自动将文档中批注提升为用户授权；用户请求处理批注后，仍按范围执行。

首个版本只需支持：读取指定或未解决线程、定位锚定段落、起草变更及回复、应用成功后回复。遇到锚点失效、相互冲突或能力不支持时保留待处理状态。Resolve 应按用户意图独立执行，不等同于已经成功修改正文。

### D. 模型协议分层，避免所有模型都承担最低兼容层成本

`tool-registry.js` 当前为兼容任意后端，采用每次一条 JSON 工具调用的模拟协议。这便于覆盖小模型，也会引入协议修复回合和串行往返。

建议保留此回退，增加统一 JSON Schema 工具定义和规范化 ToolCall/ToolResult；能力确认后使用模型原生工具调用/结构化输出。只读、无依赖的查询可批量执行，Word 写入仍受统一调度。协议升级是否节省时延和成本应通过用量记录验证。

### E. 统一上下文预算与可恢复状态

`conversation-history.js` 按估计 token 预算保留近期回合，并固定预留 8192 token；问答可能又附加整篇文档。这个组合不是覆盖所有输入和输出的模型窗口预算。工具循环虽有额外字符/传输上限，也不替代模型 token 预算。

建议按实际模型窗口统一分配系统指令、历史摘要、当前请求、已读证据、工具结果与输出。压缩时保留目标、未完成任务、用户约束、提案决定和来源指针，并可重新读取原始证据；不把摘要当成新的事实来源。

聊天目前主要使用 localStorage，存在约 4 MB 总软上限；附件已经有 IndexedDB 基础。应将会话、ChangeSet、任务事件和检查点迁往版本化 IndexedDB，保留旧数据迁移。恢复后重新验证文档身份和基线，不自动重放状态不确定的 Word 写入。

### F. 明确权限、密钥与实际数据去向

`app-state.js` 持久化配置包含 API key；前端 CSP 为自定义服务允许较宽的网络连接。这些是开放接入的代价。Claude 自身也在 WebView 存储凭据，因此没有依据声称对方一定使用系统钥匙串。[Claude 凭据说明](https://claude.com/docs/office-agents/data-storage)

Claric 可增加仅本次会话密钥模式；企业部署使用短期身份令牌/服务端代理。若使用本地系统钥匙串，必须通过受限且有鉴权的本地伴随服务实现，网页不能凭空获得钥匙串能力。将 MCP 注解作为风险提示而非可信授权，继续保留真实工具调用的审批与审计。

产品应显示本次将发送的内容类别及目的地；支持管理员限定模型、工具、域名和自动应用范围。真正的本地模式还需要处理前端资源、Office.js/宿主依赖与网络出口，单独选择 Ollama 不足以证明完全离线。

## 6. 建议实施顺序与验收

| 顺序 | 改进 | 首次交付边界 | 验收证据 |
| --- | --- | --- | --- |
| 1 | 文档证据与引用导航 | 先覆盖问答和正文 proposal 的块定位 | 每个引用对应已读证据；重复文本不误跳；过期位置明确报错 |
| 2 | 批注处理闭环 | 指定/未解决线程的修改与回复草稿 | 原生锚点正确；写入失败不回复成功；用户可逐项审阅 |
| 3 | 统一 ChangeSet 与 writer | 接入正文、格式，再逐步表格/图片 | 重叠变更可解释；冲突拒绝；部分写入可识别；连续任务减少中断 |
| 4 | 上下文与模型协议 | 全请求预算、语义摘要、原生工具调用可选 | 不超模型窗口；事实可追溯；记录协议错误率、回合数和用量 |
| 5 | 持久化与恢复 | IndexedDB 迁移和任务检查点 | 重开历史可用；旧数据保留；不重复应用已尝试的写入 |
| 6 | 企业部署与策略 | 短期令牌、网络/工具策略、审计出口 | 密钥不进入导出和日志；策略生效；可解释实际数据流 |
| 7 | 跨应用与模板深度 | 复用现有协调协议逐个增加宿主 | 先有业务需求与宿主能力验证，再扩展 Excel/PPT/Outlook |

前三项是功能价值优先顺序；凭据和数据出口的硬化可在相应部署交付时同步完成。若马上进入企业分发，第六项应前移为发布条件。

推荐的目标链路是：

```text
用户请求 → 已有能力规划 / Task Graph
                    ↓
        公共文档上下文与证据定位
                    ↓
        各专用执行器 + 共享草稿 overlay
                    ↓
          ChangeSet + 验证 + 审阅
                    ↓
      公共 writer → Office.js → 写后核验
                    ↓
          事件日志 / 检查点 / 结果引用
```

## 7. 对照测试建议

使用可丢弃测试文档，在同一 Word 宿主上比较。先让 Claric 使用与 Claude 插件尽可能接近的 Claude 模型，减少模型因素；再单独评价本地模型部署的成本和效果。账号模式、模型版本与可用工具不一致时应记录，不能宣称完全控制了变量。

测试集覆盖：中文长文档；重复句子与同名标题；多级编号；合并单元格与三线表；脚注、域、公式和图片；现有修订；多轮批注；生成后人工改动导致基线过期；插件关闭后的恢复。

比较任务完成率、定位正确率、无关改动、格式/结构损伤、修订可审阅性、冲突处理、取消恢复、用户中断次数、端到端时延，以及能取得的输入/输出/缓存 token。没有精确用量时标记未知，不以订阅价格推导单次真实模型成本。

推荐的产品定位是：**可选择模型和部署方式、证据可定位、修改可审阅的专业 Word 文档助手**。先将 Word 内的问答—定位—批注—修改—验证串联，再根据真实需求扩展跨 M365 工作流。
