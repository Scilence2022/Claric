# Claric 回合流架构深度分析与改造方案

> 本报告是 `docs/turn-flow-analysis.md` 的补充分析，不修改原报告。目标是解决“一次用户输入同时包含多个执行管线需求，并且能够按上下文、资源与依赖灵活处理”的架构问题。代码引用均按当前仓库行号记录；行号变化后应以符号/函数名复核。

## 1. 结论先行

Claric 已经具备一个很好的安全底座：模型只产生聊天结果或 proposal，Word 写入必须经过用户 Apply；路由、各专用 runner、Word diff/reassembler、工具循环都已有清晰的局部不变量。当前 compound 也不是空白能力：规划器将输入拆成最多 6 个任务，随后按原顺序复用既有 runner（`src/lib/task-planner.js:28-38`、`src/taskpane/conversation.js:1990-2057`）。

但 compound 目前仍是“串行调用多个旧式 turn”的薄编排层，而不是统一任务运行时。它缺少任务级资源声明、依赖图、跨任务结果引用、批量 proposal、提交事务和冲突重计划。因此它在以下场景会失去灵活性：

- “给全文加标题、润色正文、把所有表格改成三线表、最后总结”需要明确先后关系和作用域；
- 一个任务生成的标题应成为后续润色/总结的上下文，而不是仅依赖下一次重新读取 Word；
- 两个任务都修改同一段/同一表时，应合并、排序或提示冲突，而不是分别生成两张互不知情的卡；
- 文档在规划、生成和 Apply 之间被用户改变时，应按版本与资源锁进行冲突处理；
- 某个子任务失败或取消时，应保留可恢复的 DAG 状态，而不是只停止剩余循环。

推荐目标是增加一个兼容层之上的 **Task Graph Runtime（TGR）**：以统一任务模型表达意图、能力、资源、输入快照、依赖和 proposal；以资源锁和文档版本保护 Word 一致性；以 graph scheduler 控制并发；以 proposal aggregate 和 application transaction 对用户呈现；以事件日志和可恢复检查点提供历史可观测性。

## 2. 当前架构的优点

### 2.1 路由与专用能力边界清晰

`routeTurn` 是纯函数，按优先级将输入分到 skill、compound、图像、表格、编辑、QA 等管线（原报告 §5；实现位于 `src/taskpane/conversation.js:431-600`）。意图抑制规则避免了“格式”误进文本改写、“清理”进 LLM 文本管线等高风险误路由。`turnForTask` 又把 planner 类型映射回既有专用管线（`src/taskpane/conversation.js:1923-1987`），便于渐进式改造。

### 2.2 Word 写入默认 fail-closed，proposal 语义诚实

模型输出先形成卡，Apply 才写 Word；卡有 applied、rejected、warning、error、paused 等状态，部分成功不伪装成全成功（原报告 §9；`src/taskpane/conversation.js:755-814`、`src/taskpane/ui/proposal-card.js:479-530`）。文本 diff、chunk bookmark、表格 patch、格式/插图 OOXML 基线共同构成了较强的 staleness 防护。

### 2.3 取消和会话隔离已经是可复用基础设施

提交拥有 `AbortController`、epoch、sessionId，消息代理会阻止失主回合继续发卡；`isBusy` 统一检查多个 controller/flag（`src/taskpane/conversation.js:688-745`）。compound 还把一个 controller 传给所有子任务，取消会停止当前任务并跳过后继任务（`src/taskpane/conversation.js:1997-2057`）。这说明引入任务级 cancel/parent-child signal 有现实基础。

### 2.4 长文档并发和工具循环已有工程化经验

文档 chunk 采用 worker pool 并在 `Promise.allSettled` 语义下隔离 chunk 失败（`src/lib/orchestrator.js:251-386`）；工具循环对协议错误、上下文预算、重复调用、Abort 有明确处理（`src/lib/tool-loop.js:257-418`）。这些机制可直接抽象为任务执行器的底层策略，而不必重写所有 runner。

## 3. 结构缺陷与风险分级

### P0：必须在扩大 compound 前解决

1. **Word 写入没有统一提交事务。** 不同管线各自 Apply：文档 chunk 逆序逐块写入，表格/图像/格式有各自 attempted 和 partial 语义；跨卡仅有 `_anyCardApplyInFlight` 互斥，而不是按文档资源建立冲突检测。已有“每张卡必须 registerController，否则可与新回合竞态”的注释，说明统一锁是实际缺口（`src/taskpane/conversation.js:765-814`）。
2. **任务之间没有资源声明和冲突图。** planner 只返回 `{type,instruction}`，没有 scope、目标、读写集、前置条件或可并行性信息（`src/lib/task-planner.js:105-141`）。因此当前 compound 只能顺序跑，无法安全地让不相交任务并行，也无法在相交时给出可解释冲突。
3. **proposal 不是一等任务产物。** 每个子任务单独调用 `dispatchTurn` 并在同一消息上各建卡（`src/taskpane/conversation.js:2032-2043`），卡之间没有 graph 节点、依赖、统一 Apply 计划或跨卡原子性。用户很难知道“全部 Apply”是否会按正确顺序执行。

### P1：会明显限制复杂需求

1. **执行顺序是用户顺序，不是依赖顺序。** planner prompt 要求按原顺序输出（`src/lib/task-planner.js:81-88`），但“先插入标题，再按标题润色”是语义依赖，不一定等于文字顺序；当前无法表示条件、join、失败后替代路径。
2. **跨任务上下文只靠 Word 重读和有限历史。** planner 不读文档（`src/lib/task-planner.js:17-19`）；工具循环会按预算逐出旧 exchange（`src/lib/tool-loop.js:184-231`）；会话历史以字符串裁剪且历史图片不会恢复（原报告 §6）。任务生成的结构化结果、被拒绝原因、定位信息没有标准 context artifact。
3. **取消恢复粒度不统一。** compound 取消等价于“跳过剩余任务”，虽已落盘的 chunk 保留，但没有持久化 graph checkpoint、可从失败节点继续的统一入口（`src/taskpane/conversation.js:2044-2049`）。不同 runner 的 paused/retry 规则无法组合。
4. **规划失败回退会丢失用户的多意图。** plan 无效时回退为禁止 compound 的单路由（`src/taskpane/conversation.js:2016-2029`），这是安全的保守策略，但不满足“灵活 handle”：应显示缺失任务并允许修订计划，而不是静默退化。

### P2：可观测性和演进成本问题

- dispatch 是长 if/else，能力注册、资源需求、proposal/apply 事务策略分散在多个模块（`src/taskpane/conversation.js:2135-2169`）。
- status/log 主要是文本流，缺少稳定的 taskId、attemptId、resource、version、decision event；历史虽能保存卡和状态，但无法重建 DAG 的完整因果链。
- tool-loop 的 finish 缺少 summary 时接受为空字符串（`src/lib/tool-loop.js:359-364`），上层可能把协议不完整误认为 no-op；应纳入统一任务结果校验。

## 4. 推荐目标架构：Task Graph Runtime

### 4.1 统一任务模型

新增纯数据模型（建议 `src/lib/task-runtime/task-model.js`），兼容旧 `turn`，至少包括：

```js
{
  taskId, parentTurnId, type, instruction,
  scope: { kind: 'selection|paragraphs|document|table|images', selector },
  capabilities: ['word.read', 'llm.text', 'proposal.stage', 'word.write.tracked'],
  reads: [{ resourceId, version }],
  writes: [{ resourceId, mode: 'text|format|structure|image|comment' }],
  dependsOn: [],
  contextIn: [], contextOut: [],
  state: 'planned|ready|running|staged|blocked|applied|failed|cancelled',
  retryPolicy, priority, idempotencyKey
}
```

planner 输出仍允许旧 `{type,instruction}`，随后由 deterministic enricher 根据 selection facts、Word snapshot 和 capability registry 补齐 scope/reads/writes。LLM 不应直接决定安全能力；未知 type、空 instruction、超限仍按当前 `parsePlan` 丢弃或进入人工确认（`src/lib/task-planner.js:105-141`）。

### 4.2 能力与资源锁

建立 `CapabilityRegistry`：每个 runner 注册 `prepare/read/stage/apply`、所需 Word API、是否可并行、是否可撤销、是否支持 retry。建立 `ResourceKey`：

- `document:body`、`paragraph:<stable-id>`；
- `table:<identity>`、`cell:<table>/<row>/<col>`；
- `image:<identityKey>`；
- `selection:<snapshot-id>`；
- `bookmark:<name>` 与 `document:tracking-mode`。

调度采用读锁/写锁：只读 QA 可与不冲突的 snapshot 任务并行；两个任务写同一段或共享 tracking mode 时串行；不同图片/表可并行准备，但 Word Apply 默认单 writer。锁必须在 prepare 前声明，在 stage 后释放读资源或升级为 apply reservation。这样既保留文档 chunk 并发的吞吐，也避免跨任务互相覆盖。

### 4.3 依赖图和调度器

planner/规则层输出 DAG 而非仅数组：`dependsOn` 表示硬依赖，`join` 表示多个任务完成后生成总结，`fallback` 表示能力缺失时的替代节点。调度器维护 ready/running/completed/blocked 集合，按资源锁和优先级启动节点。

典型输入：

```text
insert-title ──┐
                ├─ polish-document ── summarize-result
format-tables ──┘
```

“标题插入”和“表格格式”若资源不相交可同时准备；两者的 Apply 仍由 writer 按 reservation 顺序提交；润色读取前两者的 staged virtual snapshot，确保模型看到的是将要落盘的文档状态，而非旧 Word 状态。planner 不可靠时，规则只生成保守串行图，并在 UI 标明“计划需要确认”。

### 4.4 执行与提案聚合

保留现有 runner 作为 adapters：`LegacyTurnRunnerAdapter` 把旧 runner 的消息、日志、proposal 转换为 `TaskResult`。每个节点产出：`readSnapshot`、`modelTrace`、`operations`、`proposalItems`、`warnings`、`retryCursor`。新增 `ProposalAggregate`：按 graph 顺序分组展示，项目带 taskId、依赖、资源、before/after、基线版本和可选项。

UI 仍可显示独立子卡以兼容旧会话，但增加“计划视图”和“Apply selected plan”入口。聚合器只组合 proposal，不绕过原有 card guard；老卡 Apply 仍单任务执行，新聚合 Apply 走统一 writer。

### 4.5 应用事务与冲突处理

Word.js 不提供跨多个 `Word.run` 的真正数据库事务，因此目标是 **可验证的微事务 + 补偿/继续**：

1. Apply 前获取 document revision（段落/表格/image identity、OOXML/hash、bookmark）；
2. 按 DAG 拓扑序和资源 reservation 写入，每个资源组一个 Word.run；
3. 每个微事务先 preflight，再打开 tracking，写入，sync，记录 applied operation；
4. 冲突时不盲写：标记 `conflict`，保留当前/提案/基线三方信息，提供 rebase（重读当前资源后只重生成该节点）或 skip；
5. 中止时保存 applied log、未处理 reservation 和 bookmark，状态为 paused，可从首个未完成节点继续；
6. 不支持的宿主能力仍 fail-closed，沿用当前 warning/partial 语义。

现有 `reassembler` 的逆序应用、re-anchor、截断拒绝可作为 adapter 内部策略（`src/lib/reassembler.js`；原报告 §10），而不是被新 runtime 重写。

### 4.6 跨任务上下文与可观测历史

引入 `ContextArtifact`：结构化保存 selection snapshot、document snapshot hash、proposal summary、applied operations、warnings、model-independent facts。任务只能通过 `contextIn` 引用 artifact，避免把大段全文复制进每个 prompt。对图像保留 identity/description，必要时按能力重新读取 bytes。

事件采用 append-only envelope：`turn.created`、`plan.validated`、`task.started`、`task.snapshot`、`task.proposal.staged`、`task.blocked`、`apply.operation`、`task.conflict`、`task.cancelled`、`turn.completed`。每条包含 `turnId/taskId/attemptId/sessionEpoch/timestamp/resourceKeys/version`。sessions 持久化时按预算优先保留事件摘要、graph 状态和 proposal 状态，再丢模型 token；这样恢复后可以从 checkpoint 继续，而不仅是把卡渲染成只读。

## 5. 兼容迁移路线

### M0：不改变行为的观测层

新增 taskId/turnId/attemptId 和结构化事件，包住现有 `dispatchTurn`、`runCompoundTurn`、proposal Apply。所有旧状态文案保持不变；补齐 planner 回退原因和 finish summary 校验。验收：旧测试全绿，历史可看到每个子任务因果链。

### M1：统一模型与 adapter

实现 `TaskSpec/TaskResult/ProposalAggregate`，把当前 `{type,instruction}` 转换为 TaskSpec；`turnForTask` 变为 registry lookup，但保留旧 dispatch 作为 fallback。先接入 QA、DOC_EDIT、FORMAT 三条低风险路径，再接表格/图像。

### M2：资源锁和单 writer

实现只读并行、Word writer 串行、资源冲突检测；旧单任务直接映射为单节点图。将 `_anyCardApplyInFlight` 逐步替换为 resource lock，但保留全局互斥作为兜底。对冲突只新增 warning/conflict，不改变已有安全拒绝。

### M3：聚合 proposal 与可恢复 Apply

把旧卡包装成 aggregate item，新增拓扑顺序 Apply、checkpoint、Continue、rebase。老会话没有 graph metadata 时按独立卡恢复为 legacy mode；不迁移旧模型 token 或附件内容。

### M4：智能依赖和虚拟快照

planner 可输出依赖 hints，但必须经过 deterministic resource analyzer 校验。对不相交任务启用并行准备；对有依赖任务注入 ContextArtifact/virtual snapshot。最后再允许用户在 UI 中调整顺序或批准冲突解决方案。

## 6. 测试验收矩阵

| 领域 | 验收场景 | 必须断言 |
|---|---|---|
| 规划 | 单意图、6 个任务、未知 type、空/非法 JSON | 兼容旧 route；未知项可解释丢弃；不静默丢用户意图 |
| DAG | 标题→润色、表格格式与图片操作、join 总结 | 拓扑顺序正确；无资源冲突可并行；join 只在依赖完成后运行 |
| 锁 | 同一段双写、同表 cell 与 row op、不同图片 | 冲突阻止盲写；不相交任务不被不必要串行 |
| snapshot | planner 后用户修改、Apply 前修改、bookmark 丢失 | 检测 revision/OOXML/hash；生成 conflict 或安全 skip；无误删 |
| proposal | 多节点混合成功/失败/无改动/拒绝 | 聚合卡逐项状态诚实；Apply 顺序可解释；部分成功可继续 |
| Word 事务 | tracking mode、表格结构、chunk 逆序、宿主能力缺失 | 每个微事务恢复 tracking；沿用 tracked/warning/partial 语义 |
| 取消 | planning、LLM、Word.run、节点间、Apply 中止 | Abort 传播；已落盘保留；未完成节点可 Continue；不启动后继 |
| 重试 | 超时、协议错误、冲突 rebase、失败 chunk | 仅重试安全节点/attempt；不重复已应用操作；retry cursor 持久化 |
| 上下文 | 前置任务输出供后置任务使用、历史裁剪、图片引用 | artifact 可引用；预算裁剪不破坏当前任务；图片按能力重取 |
| 恢复 | reload、旧 session、孤儿 bookmark、graph 中途崩溃 | 旧卡 legacy 可读；新图恢复状态；孤儿资源可回收；状态不伪造 |
| 观测 | 每种状态转换和资源冲突 | 事件包含稳定 ID、版本、资源；可从日志重建 DAG 与最终结果 |
| 回归 | 现有单任务全部管线与全套 Jest | 单任务行为不变；`npm test`、lint、轻量文档检查通过 |

建议先为 M0/M1 写纯函数测试（task normalization、resource overlap、topological scheduler、aggregate state reducer），再用 mock Word host 验证 M2/M3；最后保留少量真实 Office smoke test 验证 Word.run/sync 与 tracking changes。不要用端到端测试替代资源冲突和状态机单测。

## 7. 最终判断

Claric 不需要推倒重来。其真正的架构资产是专用管线的安全边界、proposal/apply 分离、Word re-anchor/diff 策略、Abort 纪律和可测试的纯工具循环。应避免继续向 `routeTurn` 和 `runCompoundTurn` 追加更多正则或 if/else；那会把“多需求”继续当作多个线性 turn，放大资源竞态和历史不可解释问题。

最小可行的关键改造是：**统一 TaskSpec + 资源读写集 + 结构化事件**；随后再加 DAG 调度、proposal 聚合和单 writer。这样既能保持现有单任务兼容，又能让一次输入中多个需求按依赖、能力、文档版本和用户选择灵活执行，并把“取消、冲突、部分成功、恢复”从分散的 runner 特例提升为系统级契约。
