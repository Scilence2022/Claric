# Claude 多文档共享：本机缓存静态分析

日期：2026-10-30。只读检查本机 Word 的公开静态资源缓存，未读取凭据数据库、会话数据库或文档正文，未连接或调用 Claude 的私有 WebSocket 服务。

## 证据来源与适用范围

从 Word WebKit NetworkCache 的资源记录中识别 `pivot.claude.ai` 的 JavaScript URL，再提取对应 `-blob` 静态响应体。找到 11 个 index 资源，其中 7 个为约 8.6–8.9 MB 主 bundle，其余为小型资源。选用缓存时间最近的主 bundle：

- URL：`https://pivot.claude.ai/m-addin/assets/index-DtB7PCI5.js`
- 缓存文件修改时间：2026-10-30 08:59:51，Asia/Shanghai。此时间不是构建时间，也不能证明当前活动窗口正在运行它。
- SHA-256：`15bf5d12f067e80ac62fbb071acba7fb3b76ab3602041ee03096a5384b6e47d6`
- 原始体积：8,880,968 bytes。
- [静态代码摘录及字符位置](/Users/song/.codex/visualizations/2026/10/30/01a18a59-bddf-7fe3-8b71-b688c70607f6/claude-cross-document-analysis/static-code-evidence.txt)
- [缓存资源清单](/Users/song/.codex/visualizations/2026/10/30/01a18a59-bddf-7fe3-8b71-b688c70607f6/claude-cross-document-analysis/inventory.json)

下文位置为原始 bundle 解码后的零起始 Unicode 字符偏移；混淆符号只适用于该构建。本文确认的是客户端已实现的代码路径，未证明本机已启用这些功能，也没有服务器实现或动态通信记录作为证据。

## 主要结论

该构建的跨 Office 文件能力内部称为 **conductor**，采用账号关联的云端 WebSocket 连接、分布在各插件实例中的本地文档代理、会话增量同步、共享文件广播，以及自然语言消息委派。

共享工作区在每个客户端组织为内存虚拟文件系统。它与本机聊天持久化数据库是不同层，也不是所有代理自动共用一个模型上下文窗口。模型可以按需读取其他代理的会话和文件，或发送消息请对方执行新工作。

## 1. 连接、身份和发现

位置：13,554 附近及 8,619,850–8,625,520。

- 基础地址：`wss://bridge.claudeusercontent.com`。
- 客户端拼接路径：`/v2/conductor/<userId>`。
- 每个模块实例生成随机 UUID，取前六位组成 `<app>-<suffix>` 的 agentId，例如示意 `word-a1b2c3`。
- 插件注册 schema 包括应用名称、任务能力说明、`receive_message`、可接收文件类型，以及显示信息。
- register 帧包含 OAuth 凭据字段；这里只识别了代码中的字段名，没有读取任何运行时值。
- status 传递文件名和 `documentUrl` 字段；此处变量来自规范化文档 URL 的 hash 函数，并非直接原始路径。`claude.fileId` 则用于另一套持久文件身份逻辑，应与临时 agentId 区分。
- 约每 5 秒检查文件 URL 是否变化，变化后更新身份信息和状态。

所以，发现单元是插件实例，而非一个应用只能有一个固定 Agent。相同应用的多个文档实例可以拥有不同 ID；实际同应用多窗口发现效果尚未动态验证。URL 中出现 userId 只能证明客户端寻址方式，不能证明服务端实际采用何种租户、组织或设备隔离策略。

## 2. 传输与消息协议

位置：2,350,130–2,356,350；2,938,100–2,944,390。

底层支持多路复用：发送时附加 `_agent_id`，接收时根据 `_for_agent_id` 定向或 `_exclude_agent_id` 分发。当前 Office 初始化路径创建一个 multiplexer 和主 transport；具备复用能力不等于每个窗口实际注册多个 Agent。

| 方向/类型 | 作用 |
| --- | --- |
| `register` / `deregister` | 注册实例及能力、注销 |
| `status` | 发布文件名等实例状态 |
| `stream` | 同步会话 messages 与 total_messages，可标记 full_sync |
| `file` | 发布 filename 与 data |
| `conductor_send_message` | 向指定 peer 发送任务/消息，可携带 files |
| `conductor_event` | 接收 connect、stream、file、status、disconnect 等事件 |
| `conductor_agent_online/offline/reset/expired` | 生命周期通知；实际离线清理由 disconnect 事件等路径处理 |
| `conductor_connected` / `conductor_replay_complete` | 注册成功与事件回放完成 |
| `clear_context` / `activate` | 上下文重置、激活请求 |

客户端状态为 `DISCONNECTED → CONNECTING → REPLAYING → LIVE`。回放期间缓存新 live 事件，完成后顺序处理；这是事件回放设计，尚不能据此声称 exactly-once 投递或完整一致性。

心跳默认约 30 秒，pong 超时约 90 秒；窗口重新可见时额外探测，5 秒未收到 pong 则关闭僵死连接。重连采用退避加抖动，鉴权拒绝有单独的退避参数。客户端发送 `peer_retention: 60`，但单位及服务器最终保留策略未核实。

## 3. 共享内容：会话流与文件，而不只是摘要

位置：2,932,290–2,934,500；7,210,890–7,212,130。

客户端构建的虚拟目录包括：

```text
/agents/<agent-id>/
  metadata.json
  status.json
  transcript.jsonl
  files/<filename>
```

`syncConversationToPeers` 使用 `lastSyncedMessageIndex` 将新会话消息作为 stream 发出；历史长度减少时可触发 full_sync。接收方将 messages 追加或重建为 JSONL。代码还过滤读取 `/agents/` 的 bash 工具调用及相应结果，以减少把已读取的 peer 内容再广播回去的重复传播。

文件可通过沙箱能力 `conductor.writeFile` 发布；其他实例通过 `readFile`、`listFiles` 读取。代码中的浏览器内 Map 和虚拟文件系统保存 peer 文件副本；这不是直接挂载真实磁盘目录，也不是直接给模型开放操作系统 shell。

模型获得三个协作工具：发现 peers、发送消息，以及对共享工作区的只读虚拟 bash。数据路径与模型上下文路径分开：程序可以读取完整共享 JSON 并直接写入本应用，而只将少量结果回传给模型。这能减少大数据在上下文中的展开，但内容已经通过共享协议传递，不能因此宣称数据未离开设备。

这与“只共享有关段落摘要”的推断有实质差别：客户端实际提供会话记录和文件的按需访问能力；它也不等于默认广播整篇 Word 文档，只有已读取并进入会话或显式发布的数据才有相应代码路径。

## 4. 委派的是本地 Agent 任务，写入仍发生在目标插件

位置：4,352,170–4,354,720；8,615,610–8,619,850。

发送侧 `send_message` 是立即返回的消息投递，并非“对方任务已经完成”。工具返回指令明确要求当前回合结束，不能假装已经收到完成回调或持续监控。

接收侧则有独立处理链：

1. 消息进入队列；同一客户端用 `isHandling` 顺序处理。
2. 等待本地 Agent 空闲，并检查文件访问策略及不可恢复上下文。
3. 校验、接收随附文件，注入来源标记与必要的协作提示。
4. 调用本地 Agent 的 `query(..., {kind: "conductor-peer", fromAgentId})`。
5. 目标 Agent 用自身 Office 工具执行任务，并可再通过 send_message 发回结果。

接收提示要求对任务回报，但纯状态确认不再回复，以防 Agent 无限互相确认。发送端无自动完成回调、目标侧可以主动发回复，这两点并不矛盾；它是多个独立回合组成的协作，而非一个同步 RPC。客户端代码中还存在接收能力差异及通用库的其他处理模式，应以具体 Office 集成路径为准。

因此，“共享一个会话”的用户体验不能直接等同为“一个中央模型拿到所有 Office API”。公开缓存呈现的是多个本地 Agent，互相查阅上下文并委派任务。

## 5. 权限分层

位置：4,343,200–4,344,580；5,838,167 附近；8,615,610 附近。

至少可识别以下层次：

- 平台 capability 和组织 `work_across_apps` 控制。
- 用户跨文件设置。
- 会话级 `conductor-workspace` 同意状态；未授权时禁止读取 peer 会话/文件及发送消息、文件和 transcript。
- 收到消息时对目标当前文件再次执行 access policy 检查。
- 目标 Agent 后续调用工具时适用的本地执行/审批策略。

还存在单回合来源固定、最多三次 unattended peer reply 的限制，不是无限制 Agent 自循环。仅凭这些客户端逻辑，无法证明服务器隔离、端到端加密、所有宿主的审批表现或数据保留边界。

## 6. 与 Claric 的直接差别及建议

| Claude 缓存实现 | Claric 当前侧重 | 设计启示 |
| --- | --- | --- |
| 会话记录、文件和状态形成共享虚拟工作区 | 显式 context request/response 和有界快照 | 增加可按需查询的证据/产物目录，减少反复人工搬运；不要默认同步完整会话 |
| 自然语言 send_message 触发目标本地 Agent | 类型化 task、proposal、review 和租约 | 可增加高层委派适配，但保留类型化任务与可验证结果 |
| 临时实例 ID、能力 schema 和状态发现 | workspace/document/instance 的显式身份 | 保留 Claric 的文档身份与版本，补充动态 capability schema |
| 事件回放、增量 stream 和浏览器内文件副本 | 协调事件、快照、恢复与写入保护 | 用现有协调协议承载增量 artifact 索引，避免新建第二套通信系统 |
| 多独立回合协作，不保证完成回调 | task/proposal 的显式终态 | Claric 应清晰区分 delivered、accepted、running、staged、applied、verified，不能把消息投递当成功 |

建议借鉴“可查询的共享证据与产物工作区”，继续使用 Claric 已有的文档版本、目标审阅与写入验证。Claude 的客户端设计证明了异步多 Agent 和共享数据层可以提升连贯体验；没有证据表明其跨文档写入具备原子事务或优于 Claric 的冲突处理。

## 尚不能落实的细节

服务端房间与租户隔离、服务器存储介质和保留时间、消息容量与实际投递保证、跨设备发现范围、重复投递处理、端到端加密、同一文档多个写入者的互斥，以及本机实际开启状态，均需进一步运行时观察或服务器资料；静态客户端代码不足以确认。

本报告对上一份架构比较补充了具体客户端证据，尤其确认了共享 transcript、虚拟工作区和本地 Agent 委派机制。它没有把缓存中存在的代码视为所有用户都已启用的线上行为。
