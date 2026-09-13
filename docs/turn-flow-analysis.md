# Claric 回合生命周期分析：从用户输入到任务完成

> 基于 main 分支源码（v1.0.2，commit 693e447）全文审读。所有结论均附 `文件:行号` 证据。
> 范围：`src/taskpane/**`（UI、路由、Word I/O）、`src/lib/**`（模型传输、编排、diff、工具循环）。
> 本文回答一个问题：**一条用户输入从敲下回车到文档/聊天呈现最终结果，会经过哪些阶段，在每个阶段可能走哪些分支，每个分支失败时用户看到什么。**

---

## 目录

1. [全景：七个阶段](#1-全景七个阶段)
2. [前置：启动与运行环境](#2-前置启动与运行环境)
3. [输入采集（点击发送之前）](#3-输入采集点击发送之前)
4. [submit()：提交管线](#4-submit-提交管线)
5. [路由：routeTurn 决策树](#5-路由routeturn-决策树)
6. [会话历史与提示构建](#6-会话历史与提示构建)
7. [十四条执行管线（逐一分支）](#7-十四条执行管线逐一分支)
8. [模型传输层](#8-模型传输层)
9. [提案卡与用户审阅](#9-提案卡与用户审阅)
10. [写回 Word：diff、对齐与 tracked changes](#10-写回-worddiff对齐与-tracked-changes)
11. [横切机制：取消、互斥、存储、安全](#11-横切机制)
12. [发现的边界情况与不一致](#12-发现的边界情况与不一致)
13. [附录：关键常量表](#13-附录关键常量表)

---

## 1. 全景：七个阶段

Claric 的核心设计约束是：**模型的输出永远不会直接写入 Word。一切写操作先暂存为"提案卡"（proposal card），用户点 Apply 才落盘，且以 tracked changes 记录。**

```
┌─①启动───────────────────────────────────────────────────────────────┐
│ Office.onReady → Word host 检查 → 设置/提示词/会话恢复 → 能力探测    │
│ (WordApi 1.3/1.4, platform) → 孤儿书签回收 → 连接探测               │
└──────────────────────────────────────────────────────────────────────┘
┌─②输入───────────────────────────────────────────────────────────────┐
│ textarea（IME 保护）→ 斜杠补全 → 选区预览 chip（200ms debounce）     │
│ → 附件校验/解析（5 个文件、10MiB）→ Enter 触发 submitCurrent()      │
└──────────────────────────────────────────────────────────────────────┘
┌─③提交 submit()───────────────────────────────────────────────────────┐
│ busy 检查 → owner{AbortController, epoch, history} → 读选区          │
│ (text+images+tableRegion，失败降级为空) → routeTurn() → 附件注入     │
│ → Proxy 包装消息句柄（会话切换/中止防护）→ dispatchTurn()           │
└──────────────────────────────────────────────────────────────────────┘
┌─④路由 routeTurn()（纯函数，17 级优先级）─────────────────────────────┐
│ /skill → 图注意图 → 复合计数≥2 → 图像管理 → 插图 → 建表 → 追加      │
│ → 格式 → 清理 → [有选区: 图像/表格对象/问句/审阅/改写]              │
│ → 文档级表格 → 文档级图像 → 编辑意图 → 问句 → 规划器兜底 → QA       │
└──────────────────────────────────────────────────────────────────────┘
┌─⑤执行管线（14 种 turn）─────────────────────────────────────────────┐
│ 每条管线 = 读文档上下文 → 构建 prompt → LLM 流式调用                 │
│   → 解析/校验（JSON 修复、delimiter 协议、think 剥离）               │
│   → 暂存 proposal → 渲染提案卡                                       │
│ 唯二例外：DOC_QA 直接流式输出到聊天；CLEANUP 无 LLM（纯 Word 扫描）  │
└──────────────────────────────────────────────────────────────────────┘
┌─⑥审阅与应用─────────────────────────────────────────────────────────┐
│ 提案卡（逐项 checkbox）→ Apply 守卫链 → apply*()                     │
│ → staleness 复核 → diff 策略链 → tracked changes 写入               │
│ → 终态：applied / rejected / warning / error / paused(可续)          │
└──────────────────────────────────────────────────────────────────────┘
┌─⑦收尾───────────────────────────────────────────────────────────────┐
│ finalizeForHistory → _commitSession → sessions.js 持久化            │
│（配额逐级降级：去预览→截断 diff→丢提案→截断文本→逐出旧会话）        │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 2. 前置：启动与运行环境

### 2.1 启动序列（taskpane.js:36-231）

整个 bootstrap 包在双重门禁内（taskpane.js:36-42）：

```js
if (typeof Office !== 'undefined') {
    Office.onReady((info) => {
        if (info.host === Office.HostType.Word) { initialize(); }
    });
}
```

**分支**：Office.js 加载失败（CDN 被拦截）或宿主非 Word → `initialize()` 永不执行，面板停留在静态 welcome 标记上，**没有任何用户可见的错误提示**（已确认的可用性缺口，见 §12）。

`initialize()` 内部严格按序执行 22 步，关键的：

| 步骤 | 位置 | 内容 | 失败分支 |
|---|---|---|---|
| 1-2 | taskpane.js:49-50 | 读 `wordAI.config`、promptManager.loadState() | 设置损坏 → "Saved settings could not be read (corrupt data). Defaults restored."（app-state.js:363），逐字段再校验（normalizeConfig，app-state.js:186-298） |
| 4 | :54-63 | 状态栏、聊天视图、citation 处理器 | — |
| 6 | :66-76 | 历史视图 | 删除失败 → "Delete session failed: …"（:73） |
| 9 | :114-130 | createConversation（含 onTurnCommitted→persistCurrentSession） | 持久化失败 → "Save session failed: …"（:88），**绝不打断回合** |
| 12-13 | :148-153 | welcome chips、模型 pill | 无模型 → pill 显示 "(no model)"（:242） |
| 14 | :155-161 | **恢复最近会话**（无会话则停留 welcome 页——这就是首次运行路径，无向导） | readSession 无 messages 数组 → null → 走 welcome |
| 16 | :168-172 | `reapOrphanChunkBookmarks` — 回收上次崩溃/重载残留的 `_wdp*` 书签 | 宿主无 `Range.getBookmarks` 或文档保护 → 静默跳过（reassembler.js:644-648） |
| 17-19 | :175-207 | Word API 版本探测（1.8→1.1 循环）、`supportsComments`(1.4)、`supportsTables`(1.3)、`platform` | 缺 1.4 → "Comment features unavailable (requires Word API 1.4)"（:197-199）；缺 1.3 → 表格同型日志（:200-202）；platform 'unknown' fail-closed（platform.js:44-47） |
| 20 | :210 | 后台 `testConnection()` 探测 | 结果只反映在连接指示灯（status-bar.js:153-178） |

### 2.2 会话恢复

- 启动：自动加载最新会话（taskpane.js:157-161）。
- 历史面板：点行 → `conversation.newChat(); chatView.setCurrentSession(session)`（taskpane.js:67-71）。
- 恢复后的提案卡由 `renderStaticProposalCard` 重建为**只读**（chat-view.js:487-562），终态徽章取自持久化的 `meta.state`；恢复的 citation pill 只在 `_citationSelectHandler` 已注册时渲染，"宁可省略也不渲染成死按钮"（chat-view.js:54-64）。

---

## 3. 输入采集（点击发送之前）

### 3.1 键盘与 IME（ui/input-bar.js:337-385）

- **IME 保护优先**：`if (composing || e.isComposing || e.keyCode === 229) return;`（:338）——中文输入法组词阶段的 Enter 不会误发。
- **Enter = 发送，Shift+Enter = 换行**（:381-384）。
- **↑/↓ 输入历史**：仅补全器关闭时生效，且用光标行位置守卫（↑ 只在第一行、↓ 只在最后一行），多行编辑时光标移动不受劫持（:353-380）；上限 `MAX_INPUT_HISTORY = 100`，连续重复折叠（:81-85）。

### 3.2 斜杠补全（:323-348, :88-130）

输入以 `/` 开头 → 过滤器 = 首个空格前的 token → 匹配 slash 或 skill 名（大小写不敏感）。打开时：↑↓ 循环移动、**Tab 或 Enter 选中高亮项**、Esc 关闭、点外关闭。选中后文本框替换为 `` `${skill.slash} ` ``。候选源 = `listSkills()`：**16 个内建 skill + 保留 `/mcp` + 每个用户保存的自定义提示词 + 导入的 SKILL.md 包**（skills.js:271-294；内建优先 `find` 匹配，同名自定义/导入技能被内建遮蔽）。

### 3.3 选区监视（word-actions.js:971-1005 → input-bar.js:540-590）

- `DocumentSelectionChanged` 事件 → **200ms 尾随 debounce**（拖拽选择只触发一次 Word.run），注册时立即发一次初始状态；Office 事件缺失 → no-op unsubscribe；注册失败被吞（预览静止，:989）。
- 预览 chip 内容：文本片段、**表格角标 `Table R1C1 → R3C2`**（多单元格区域，优先于图像显示）、图片缩略图（`W×Hpt` 工具提示）、截断时 `+N` 徽标（单次快照上限 `MAX_SELECTION_IMAGES = 6`，word-actions.js:839）。
- 纯图片选区 `selection.text === ''` 也能被识别——这是"选区即对象"模式的基础（word-actions.js:937）。

### 3.4 附件（input-bar.js:247-303 + lib/file-attachments.js）

限制：**5 个文件、单文本 10MiB、单图 4.5MiB、总 10MiB、上下文 200K 字符**（file-attachments.js:31-37）。

流程分支（每文件独立）：
1. `detectAttachmentKind`：扩展名优先、MIME 兜底（".md 以 octet-stream 送达仍是文本"，:44-46）；都不中 → "unsupported file type (use text, image, .docx or .pdf)"。
2. `validateAttachment` 顺序拒绝：未知类型 → 超单文件上限 → 超 5 个 → 超总量；返回 `{ok:false, error}`，**不抛异常**。
3. 通过 → **立即挂 pending chip**（spinner + "Parsing…"）——文件选中的瞬间就有反馈（commit 693e447 的行为）。
4. 顺序解析：text → `file.text()`/FileReader；image → 分块 btoa（0x8000，防 `String.fromCharCode` 栈溢出，:246-249）；docx → 动态 import mammoth `extractRawText`；pdf → 动态 import pdfjs legacy 构建，逐页 `streamTextContent`，双层 finally 清理且"清理不得吞掉结果或错误"（:313-329）。任何解析错误 → 移除 chip + 记录错误，**其余文件继续**。
5. 发送时仍在解析 → submitCurrent 拒绝："Attachments are still loading. Wait before sending."（input-bar.js:172-175）。
6. **持久化刻意排除内容**：会话存储只留 `{name, kind, size}` 元数据——"提取文本和图片 data URL 永不进入 localStorage（约 5MB 上限）"（message-shape.js:41 注释）。

### 3.5 发送/取消按钮与门禁汇总

`setProcessing(is)` 把发送键 `↑` 变为 `■`，同时**禁用 textarea 和附件按钮**（input-bar.js:502-510）。点击 = 处理中则 `onCancel()`（→ `conversation.cancel()`），否则 `submitCurrent()`。

发送门禁（按时序）：
1. 附件解析中 → 拒绝（:172-175）。
2. 空文本且无附件 → no-op（:176）。
3. `conversation.submit` 内 `isBusy()` → "Already processing. Cancel the current run first."（conversation.js:2190-2193）。
4. 路由后 turn 为 null（空输入）→ 静默返回（:2253）。

无任何针对"后端未配置"的提交门禁——设置不完整只表现为连接指示灯、"(no model)" pill 和回合内的错误消息（settings-loader.js 无提交校验）。

---

## 4. submit()：提交管线

`submit(text, attachments)`（conversation.js:2182-2302）：

1. **附件兜底**：空文本 + 有附件 → 有效文本设为 `'What do the attached file(s) say?'`，保证附件-only 提交路由到 DOC_QA 而非规划器（:2188）。
2. **领取回合**：`owner = {controller: new AbortController(), epoch, sessionId, conversationHistory}`；`submissionOwner = owner`；置 `isProcessing`、锁输入（:2195-2204）。`isCurrentSession()` 同时校验 epoch 和 sessionId——切换会话或新聊天都会使在途回合"失主"。
3. **读选区**（:2229-2240）：`getSelection()`（= `readSelectionContent`）返回 `{text, images, hasMultiCellTableRegion, …}`；**任何读取失败 → 全部降级为空**（当作无选区继续，:2236-2240）。图像只留元数据 `{width,height,altText,identityKey}`，base64 不进路由层（`_normalizeSelection`，:2348-2362）。
4. **路由**：`routeTurn(effective, {hasSelection, hasImageSelection, hasTextSelection, hasMultiCellTableRegion, skills})`（:2246-2252）。`hasSelection = 文本 || 图片 || 多单元格区域` 三者任一。
5. **附件注入**（:2258-2265）：文本类 → `buildAttachmentContext` 生成 `--- ATTACHED FILE: name ---` 块追加到 `turn.question`/`turn.instruction`；图片 → `turn.questionImages`（仅 QA 管线实际作为 image_url parts 发送；其他管线只在上下文块中列名）。
6. **消息句柄 Proxy**（:2272-2283）：所有 msg 方法调用都先检查 `isCurrentSession()`，且中止后禁止 `attachProposal`——**中止的回合不能再往聊天里塞卡**。
7. **dispatchTurn** → 对应管线（§7）。finally：折叠日志/模型输出，若中止 → "Cancelled."，`finalizeForHistory()`，`_commitSession()`（:2293-2301）。
8. **收尾释放**（finally of submit，:2207-2217）：仅当 owner 仍是当前 submissionOwner 才清 busy 标志——防止打断后继者的锁。

**回合级错误呈现**（`_reportTurnError`，:755-763）三分支：
- `AbortError` → 状态 "Cancelled."（不是错误）；
- `error.noChanges` → 状态显示其 message（"模型没改动"是状态不是错误）；
- 其他 → `markError`（红字 "Error: …"）。

---

## 5. 路由：routeTurn 决策树

`routeTurn` 是纯函数（conversation.js:431-600），评估顺序即优先级，**每级命中即返回**：

| # | 条件 | 结果 | 证据 |
|---|---|---|---|
| 1 | 空输入 | `null` | :439-440 |
| 2 | `/skill args` 精确匹配首 token | `SKILL` | :451-454 |
| 3 | 图注/图例意图 &&（选了图片 \|\| 无选区且非问句）&& 非表格区域 | `IMAGE_TOOL` | :462-471 |
| 4 | 意图族计数 ≥ 2（见 5.4） | `COMPOUND` | :477-491 |
| 5 | 图像管理意图（删/换/缩放/alt/多图设计） | `IMAGE_TOOL` | :496-504 |
| 6 | 插图意图（插图/配图/svg/设计示意图…） | `ILLUSTRATION` | :508-510 |
| 7 | 建表意图（创建动词+表格 / N行N列 / 3x3） | `TABLE` | :515-517 |
| 8 | 追加意图（续写/追加/继续写/to the document…） | `DOC_APPEND` | :520-522 |
| 9 | 格式意图（加粗/居中/标题/heading…）：若多单元格区域 && 表格样式词（边框/底纹/三线表…）→ `TABLE_TOOL`；否则 `FORMAT{scope}` | :525-534 | |
| 10 | 清理意图（删除空段落/delete empty paragraphs） | `CLEANUP` | :537-539 |
| 11 | 有选区：纯图片→`IMAGE_TOOL`；多单元格→`TABLE_TOOL`；问句→`DOC_QA`；审阅意图无编辑动词→`DOC_QA`；否则 `SELECTION_EDIT` | :540-570 | |
| 12 | 文档级表格意图（"所有表格…"，复数标记） | `DOCUMENT_TABLE_TOOL` | :574-576 |
| 13 | 文档级图像意图（"所有图片…"） | `DOCUMENT_IMAGE_TOOL` | :577-579 |
| 14 | 编辑意图（edit/polish/润色/修订…） | `DOC_EDIT` | :580-582 |
| 15 | 问句（疑问词开头/吗呢/问号结尾） | `DOC_QA` | :584-586 |
| 16 | 允许 compound | `COMPOUND`（规划器分类模糊指令） | :589-598 |
| 17 | 规划器被禁用 | `DOC_QA` 兜底 | :599 |

### 5.2 意图识别的关键构成（全部中英双语正则）

- `EDIT_INTENT_RE`（:69）：英文词边界动词 + 中文子串；**更新/充实/增补/扩写类动词必须邻接文档类宾语**——"是谁更新的"不会误入编辑管线。
- `QUESTION_LEAD_RE`（:76）：仅匹配**开头**；配合 `[吗呢]` 与尾问号检测（:89-95）——"你能改变选择的表格的样式吗？"这类主语开头的中文是非问句因此落入 QA 而非 FORMAT（后者曾以"no changes"收场的 bug 已修）。
- `APPEND_INTENT_RE`（:116）：英文要求显式目标（"to the document/end"），避免裸 "add a paragraph" 落入追加。
- `ILLUSTRATION_INTENT_RE`（:132）：插图/配图/svg 直配；**泛化名词（图片/图像/示意图/流程图）必须邻接创作动词**——"修改图像描述的措辞" stays edit。
- `TABLE_INTENT_RE`（:271）：创建动词+表格、显式维度（含中文数字）、NxN；裸"表格…行"不算。
- `IMAGE_TOOL_INTENT_RE`（:171）：管理动词+图片名词双向窗口；多图数量词（两/三/几张）；`set … alt`。
- `DOCUMENT_*_INTENT_RE`（:194,:201）：必须有复数标记（所有/全部/每个/all/every…），单图"把图片居中"仍走 IMAGE_TOOL。
- `CLEANUP_INTENT_RE`（:324）：删/清除 + 空段落/空行（中英）。
- `REVIEW_INTENT_RE`（:345）：check/review/检查/分析——**仅在选区分支使用**，且 `!looksLikeEditIntent` 才转 QA（"检查并修改这段话"是编辑）。

### 5.3 互斥与抑制规则

- **问句压倒一切意图**：每个 `looksLike*Intent` 第一行都是 `if (looksLikeQuestion(text)) return false;`（如 :104-107）——"如何给文章配插图？"是 QA。唯一例外是图注分支显式自查（:462）。
- **表格抑制追加**：`countIntentFamilies` 中 `if (!table && looksLikeAppendIntent)`（:375-377）——"到文档末尾插入一个表格"的"到文档末尾"是位置短语不是追加请求。
- **插图 > 追加/编辑**、**建表 > 追加/选区**（表格可锚定选区 before/after 插入）：见 :505-517 注释。
- **格式不改写文本**：格式意图必须先于选区分支（:523-525），否则"给选中文字加粗"会进入文本 diff 管线重写内容。
- **清理永不进文本管线**：解析器从不产出空段落（document-parser.js:190），对齐器也排除它们——文本管线**结构上不可能**服务此意图（conversation.js:535-536）。

### 5.4 复合计数（:477-491）

```
families = countIntentFamilies(text)            // 插图/表格/追加/格式/编辑/清理 六族
         + docCompound      (!imageSelected && 文档图像意图 ? 1 : 0)
         + tableDocCompound (!multiCellRegion && 文档表格意图 ? 1 : 0)
>= 2 → COMPOUND
```

### 5.5 规划器任务的再映射（`turnForTask`，:1923-1987）

- **清理拦截**：任何非 qa 任务的指令若匹配清理意图 → CLEANUP，**无视规划器给的类型标签**（通常是 edit）——规划器没见过空段落概念（:1932-1934）。
- `insert` → FORMAT{scope:'document'}（结构性插入如标题放文档域，插进选区会放错位置）。
- `edit` 三叉：选图+图注意图 → IMAGE_TOOL；有文本选区 → SELECTION_EDIT；纯图选区 → IMAGE_TOOL；否则 DOC_EDIT。
- `image_management`：有选中图 → IMAGE_TOOL（选区锚定）；否则 DOCUMENT_IMAGE_TOOL（全文档快照）。
- `table_management` → DOCUMENT_TABLE_TOOL（全部表格一个会话）。
- 未知类型 → DOC_QA（switch default）。

---

## 6. 会话历史与提示构建

### 6.1 会话历史（lib/conversation-history.js）

- **进入每个请求**：`submit` 时从持久化 session 构建 `buildConversationHistory(records)`（conversation.js:2197），经 `withConversationHistory` 前置到所有请求消息（6 个调用点：orchestrator.js:278 每 chunk 一次、comment-request.js:96、word-actions.js:69/:3745、agent-actions.js:456/:464）。
- **保留内容**：用户文本、附件**名字**（注明字节已不可用）、**未出错未取消**回合的助手文本、失败/取消回合的真实状态行、提案标题+状态注释+逐项 before/after。状态措辞防止模型把 proposed 当 applied（conversation-history.js:184-190）。
- **上限**：每字段 1000 字符、提案 8 条、项 8 条、附件 8 个（:27-30）；token 预算 = `contextBudgetTokens`（默认 128000，钳到 1000…2M）− 8192 保留（:58-67）。
- **裁剪**：按整回合（user 开头分组）保留**连续后缀**，最旧先丢；**最新回合永不丢弃**——超大最新回合改为从最新消息起截断并标 ` [trimmed]`（:90-93 注释明确指出这正是本模块要修的连续性 bug）。CJK ≈ 1 token/字，其余 ≈ 4 字符/token（:40-48）。
- 历史条目一律 string content——**历史图片永远不会重新进入请求**。

### 6.2 提示词三来源与优先级（lib/prompt-manager.js）

1. `templateOverride`（每次调用参数，`!== undefined` 即生效——空串覆盖激活提示词，:336-338）；
2. localStorage 激活提示词（`wordAI.prompts.{category}` / `wordAI.active.{category}`，:260-276）；
3. 无 → 返回 `[]`（上游决定降级）。

种子模板来自 `prompts.json`（legal-review / plain-english / risk-analysis），由外层加载进 PromptManager——PromptManager 自身不读该文件。激活互斥：激活 amendment 清除 comment，反之亦然；`getActiveMode` 返回 summary > amendment > comment，**'both' 已不可达**（:244 注释）。

### 6.3 占位符替换与协议防伪

- `{selection}`：全局替换；模板无占位符则 `\n\n` 追加（永远发送，:347-353）。
- `{comments}`：同上追加降级。
- `{whole document}` / `{tracked changes}`：**仅在占位符存在且数据提供时替换，无追加降级**——调用方漏传时字面量静默保留（:475-482，见 §12）。
- **defangProtocolMarkers**（response-parser.js:103-120）：把 `===AMENDMENT===` 等协议标记的字母核内插 U+200B（在字母中间断开，而非第 0 字符后——后者可被变长 `=` 重组成完整标记）。文档文本/选区文本进 prompt 前一律 defang；**解析先于 restore**，否则文档里回显的标记会制造假节边界（orchestrator.js:312-316）。

---

## 7. 十四条执行管线（逐一分支）

`dispatchTurn`（conversation.js:2135-2171）把 turn 映射到 runner。以下每条管线按"准备 → 调用 → 解析 → 暂存 → 应用"展开，分支表列出全部已验证路径。

### 7.0 通用骨架

所有 runner 共享：
- `_beginChatTurn(turnController)`（:731-737）：领取/复用 AbortController、置 busy、锁输入。compound 回合传入共享 controller，一个 cancel 停全链。
- `actionDepsFor`（:688-707）：日志同时写入状态栏与消息的折叠工作日志；`logWithRetry` 的 retry 闭包自带 busy/会话守卫。
- 每个 LLM 调用带 `signal`；每个 `checkOperationSignal` 立刻抛 `AbortError`。

### 7.1 DOC_QA（问答，conversation.js:1761-1786 + word-actions.js:3690-3774）

**准备/组装**（word-actions.js:3692-3737，按序拼接）：
1. context 激活提示词（仅当**无会话历史**时作为前缀；有历史时改为以 system 消息身份进入请求，:3746-3747）。
2. skill 模板（若有）。
3. 问题本体 + 附件上下文块（submit 已注入）。
4. 上下文三叉（互斥）：
   - 有文本选区：`readSelectionTableContext` 优先（表格选区读成 markdown 网格并标注覆盖区域；混合选区按段落+网格）→ 失败返回 null → 降级 `--- SELECTED TEXT ---`；有图元数据则附 `--- SELECTED IMAGES ---` 对象引用块（**字节永不注入**）。
   - 无文本但有图元数据（防御分支，规划器 qa 任务可能落到这）→ 仅图像引用块（:3723-3727）。
   - 裸光标 → `readCursorContext`：向前走最多 120 段找最近标题注入 `--- CURSOR LOCATION ---`（:3728-3736）。
5. `--- DOCUMENT ---` + 全文（按设置 richness：plain/headings/structured）。

**调用**：流式，**300s 空闲超时**（比客户端默认 120s 长，:3742-3743 注释）。

**分支**：
| 情形 | 行为 |
|---|---|
| 有上传图片附件 | 组装 `[{type:'text'}, …{type:'image_url'}]` parts；HTTP 4xx（纯文本后端）→ 日志警告 + **去掉图片重发一次**（:3762-3768）；Abort/Timeout/非 4xx → 原样抛出 |
| 有会话历史 | `sendMessagesStream`（保留角色） |
| 无历史 | `sendPromptStream`（扁平 prompt；reasoning 通道在此被丢弃） |
| 流被截断（无 [DONE]/finish_reason）| 客户端抛错 → markError |
| finish_reason=length | "LLM output truncated (finish_reason=length)…Reduce the scope…"（llm-client.js:386-393） |

成功后 `msg.setText(answer)` 重渲染（流式期间原始 token 含 think 标签，最终剥离）。

### 7.2 SELECTION_EDIT（选区改写，conversation.js:1021-1151）

**路由前分流**（word-actions.js 读侧）：
- 多单元格表格选区 → `_prepareTableAmendment`（坐标网格 R1C1 + JSON patch 协议）。
- 混合选区（段落+表格，如题注+表+注）→ `readMixedTableSelection` 段落粒度（:517-547）。
- 其余 → 平文本。

**执行序**（conversation.js:1032-1067）：
1. **链式指令**（含 然后/接着/并且/，再 等，CHAIN_RE :236）→ 直接走 `prepareTableToolEdit` 工具循环；非表格选区返回 null 落回单发。
2. 单发 `prepareSelectionAmendment`：`composeMessages('amendment')`（或 merged 模式 `composeMergedMessages`）→ 流式 → `parseDelimitedResponse`（merged 时）。
3. **表格 patch 解析失败**（`/no JSON object/i`）→ 日志警告 + **改走工具循环重试一次**（:1054-1062）——逐步验证挽回一次性 JSON 协议失败。工具循环仍无果 → 原错误抛出。
4. merged 模式两个标记都缺 → **二次分类调用**；仍失败 → 整个原始响应当 amendment（:1082-1091）；**二次调用中的取消必须传播**（:1095 `if (fallbackError.name === 'AbortError') throw`）。
5. `proposal.noOps`（工具循环只读收场）→ 直接把 summary 设为聊天答案，**不建卡**（:1072-1076）——否则 Apply 会谎报 "Applied as tracked changes"。
6. 表格 patch 且 0 项 → "The model proposed no changes."（:1082-1085）。

**暂存卡**：平文本 = 单项 before/after diff；表格 = 每 cell/rowOp/merge/styleOp 一个 checkbox（`filterTablePatchBySelection` 按卡片序 cells→rowOps→merges→styleOps 过滤，:392-404）。

**应用**（`applySelectionAmendment`，word-actions.js:1237-1341）——先表格路线、再混合路线、再平文本：

平文本（一个 Word.run）：
| 分支 | 级别 | 用户所见 |
|---|---|---|
| 选区文本与暂存时不一致（标准化比较）| **跳过写** | 卡警告 "Selection changed since this proposal was staged…"（:1269-1272 → conversation.js:1124-1127） |
| 粒度 diff 抛错 | 降级 | 整选区替换（先**重新确认 tracking 模式**——失败策略可能把模式留在任何状态，:1294-1303） |
| 替换也抛错 | 致命 | 卡 "Apply failed: …" |
| 注释插入失败 / 宿主无 1.4 | 警告 | 注释文本记入日志不丢失（:1336-1341） |
| finally | — | tracking 强制 off（注释等后续操作不得继承）|

表格 patch（:1375-1657）五阶段**顺序是正确性约束**（:1395-1398）：①单元格文本（原坐标仍有效）→ ②区域绑定样式 op → ③行结构 op（降序执行；`insertRows`/`TableCell.merge` 缺失则警告跳过；非桌面宿主行 op **不追踪**并警告）→ ④合并（永远不追踪）→ ⑤表级样式。预检全部致命：引用的 tableIndex 不存在（多表）、选区已不在表中（单表）、行数变化、**任一被改单元格/行的现文本与 originals 不符**（列出至多 3 个坐标，:1479-1508）。返回诚实计数 `{cellsApplied, cellsSkipped, …, warnings}`；0 应用或带警告 → 卡 `markWarning`。

混合选区（:1994-2119）：段落对齐（复用 reassembler `_alignParagraphs`）；截断守卫（<30%）为**普通 Error 直接上卡**（:2038-2042）；**刻意没有整选区替换兜底**——该兜底恰恰会摧毁表格（:1985-1988）。表格段落的 delete/insert 一律跳过警告。

### 7.3 DOC_EDIT / 文档级 amendment（chunked 管线）

入口 `runDocumentTurn`（conversation.js:821-876）→ `runDocumentSkill(gateApply:true)`（word-actions.js:3264-3411）：

```
parseDocument（4 次 sync；空白段被排除）
→ chunkDocument({maxTokens: 6000})        ← 表格段被跳过、chunk 不跨表（含合并屏障）
→ extractContext（定义 + 标题大纲）
→ bookmarkChunkRanges（隐藏 _wdp* 书签；先校验边界段文本未漂移，漂移=致命）
→ processChunksParallel（并发 6；任一 chunk >8000 tokens 则 4；每 chunk 300s）
→ [gateApply] 返回 {staged, apply, discard, retryFailed, …}
```

**每 chunk 的分支**（orchestrator.js:263-383）：
| 情形 | 结果 |
|---|---|
| 提前中止 | status 'cancelled'，不计时 |
| `chunk.oversized`（单段超预算，无法按段书签）| **网络调用前**即抛，提示拆分段落 |
| TimeoutError / 其他错误 | status 'rejected' + 消息；**不阻断其他 chunk**（Promise.allSettled 语义） |
| merged 模式响应无任何 delimiter | **拒绝**——"错误的文本恰好在正确的长度上"，30% 截断守卫抓不住它（:319-330） |
| 空响应 | 记 fulfilled 无修改（不是错误，:337-342） |
| 有 onChunkToken | 流式 + reasoning 捕获 |

**暂存**（`stageDocumentProposal`，conversation.js:883-1016）：只保留 amendment 与原文不同的 chunk（归一化比较，:886-888）。分支：
- 有失败且有成功 → 日志挂 **Retry 链接**（`retryFailed` 保留失败 chunk 书签以重驱动，word-actions.js:3342-3350）。
- 全失败 → 状态行给出真实原因 + 首个错误 + Retry 链接（:894-921）；**没有卡**（retryProposal 时注释-only 结果仍可成卡）。
- 全部"无修改" → `discard()` 清书签 + "The model proposed no changes."（:923-927）。
- 正常 → 卡：每 chunk 一个 checkbox（内联 diff + §定位链接）。

**应用**（`applyChunkResults`，reassembler.js:841-1086）：
- 修正按**文档逆序**应用（长度变化不使更早范围失效），每 chunk 独立 Word.run，chunk 间 yield 事件循环。
- 每 chunk：解析书签 → 丢失则**跳过**（'bookmark range lost'）→ **re-anchor**（在当前范围内定位存储的段落窗口——用户先用另一张卡插入了标题也不会误删漂移段；失败跳过不冒险）→ 段落级策略 → 失败降级范围级策略 → 也失败则该 chunk 记错。
- `TruncatedOutputError`（<30%）**刻意无兜底**——范围级策略只会把同样截断的文本写得更糟（reassembler.js:23-30, :937-944）。
- 协作暂停：每个 chunk 边界检查 signal；中止 → `interrupted`，**剩余书签保留**，卡变 "Continue applying"（conversation.js:957-967）。
- 注释阶段：文档顺序；插入失败视为 context 已污染 → 弹出重驱动余项（shrink-guaranteed 循环，:1034-1083）。
- 完成后清理书签（失败 chunk 的**刻意保留**供 retry；中断时全部保留）。

### 7.4 DOC_APPEND（追加，conversation.js:1158-1203）

`prepareDocumentAppend`（word-actions.js:2140，纯文本规则 prompt）→ 空生成文本 → "The model returned no content to append."；否则卡（0 → N chars）；Apply 按 `\n+` 拆段逐段 `body.insertParagraph(…, end)`（:2183-2207），无锚点失效问题。

### 7.5 FORMAT（格式 op，conversation.js:1212-1282）

`prepareFormatProposal`（word-actions.js:2561）：宿主无书签 API → 拒绝；选区漂移 → 拒绝；**先插 `_claric_fmt_*` 锚点书签 + text/OOXML 基线，再调 LLM**；空 ops 或出错 → 立即弃锚。LLM 返回严格 allowlist 的 JSON op 数组（font/paragraph/insert）。卡：每 op 一个 checkbox + locate。

Apply（:2620-2695）预检全致命：无 ops、锚点缺失、**已尝试过**（一次性守卫）、书签丢失/文本漂移、**OOXML 基线不匹配**（抓住纯格式漂移）。运行中：`insert` op 逆序入队保序；match 搜索串截 255（Word 上限）；列表 op 需 WordApi 1.3 特性检测（缺失 → 警告跳过，:2837-2840）；**每属性独立 try/caught**（单属性失败继续批次）；零目标命中 → 卡警告 "Nothing applied — no formatting targets matched."；首写之后出错 → `partial=true` + 卡警告（不重抛）。`attempted=true` 在首写前置位；8 处信号检查让 Stop 落在 op 之间。

### 7.6 TABLE（新建表，conversation.js:1293-1356）

`supportsTables === false` → "This Word host does not support the table APIs (WordApi 1.3)…"（:1294-1296）。`prepareTableProposal`（word-actions.js:2240）：
- 显式维度且无内容措辞（`TABLE_CONTENT_HINT_RE`）→ **确定性空网格，零 LLM 调用**（model:null）。
- 有内容/无维度 → 严格 JSON 契约；指定维度则复述为硬约束，**网格不匹配整体拒收**（:2295-2304）。
- spec 缺失 → "No table could be drafted from this instruction."

Apply（:2328-2388）：**spec 在应用时重新校验**（提案可能经会话持久化往返，视为不可信）；桌面才追踪插入，否则警告 "Table insertion cannot be tracked … applied directly."；`Word.ChangeTrackingMode` 整体缺失 → 追加专门警告；styleBuiltIn/autoFitWindow 特性检测。

### 7.7 CLEANUP（清理，conversation.js:1654-1696）

**无 LLM**。`prepareEmptyParagraphCleanup`：Word.run 扫描——仅空白文本 && 非文末段（Word 要求尾段标记）&& 非表内 && **无内嵌图片**（图片住在"空"段里，删段即毁图，:2399-2400）。0 个 → "No empty paragraphs found."。卡 "Delete N empty paragraph(s)"；Apply **再次扫描**（不信暂存计数），逆序删除；0 个 → 卡警告 "Nothing applied — no empty paragraphs remained."

### 7.8 ILLUSTRATION（插图，conversation.js:1365-1430 + word-actions.js:2926-3120）

**光标位置才建锚**（`_claric_img_*` + OOXML 基线，**基线缺失即删锚抛错**，:2945-2949）。

**渲染器选择**（`illustrationRenderer`，illustration.js:404-407，调用点 word-actions.js:2979）：
1. 显式矢量措辞（`\bsvg\b|矢量图?|…\bvector\b`）→ `'svg'`（图像模型产不出 markup）。
2. 否则图像模型就绪（`getActiveImageConfig` 非 null）→ `'image'`。
3. 否则 → `'svg'`（旧安装行为不变）。

**image 路线**（word-actions.js:2981-3027）：
1. 有会话历史 → 先用**聊天模型**把"那个建议的第二张"解析成自足图像 brief；`UNRESOLVED` 或空 → 抛 "The referenced image details could not be resolved…"（不猜）。
2. `generateImage`（image-client.js:389，180s）：校验链（url/model/prompt/端点合法性 HTTPS|localhost|相对代理）→ MiniMax `/image_generation`（aspect_ratio+base64）或其余 `/images/generations`（`gpt-image-*` 不发 b64_json 字段否则被拒）→ 响应归一化（b64_json / data:URL / hosted URL 下载——下载 URL 仅 HTTPS/loopback，CORS 失败信息明说可能原因）→ 8M base64 字符上限。
3. **失败 → Abort 重抛；其余 → 日志警告 + 落回 SVG 路线**（:3022-3026）。

**svg 路线**：`buildIllustrationPrompt`（新图禁止文字/外链/script）或 redesign prompt（四级保真；4xx 且要求视觉输入且无源 SVG → 拒绝；否则剥离图像重试一次）→ `parseIllustration`（剥 fence、截 `<svg>…</svg>`，>256KB → null）→ `sanitizeSvg`（DOMPurify svg+svgFilters profile，禁 foreignObject/iframe/script）→ `ensureSvgDimensions`（缺失注入 1200×800，防 ~300×150 默认栅格模糊）。sanitizer 清空 → 统一报 "The model produced no usable illustration."（无专门消息，§12）。

**Apply**（word-actions.js:3067-3120）：SVG 先栅格化 PNG（`insertInlinePictureFromBase64` 不收 SVG）；cursor 位置三重校验（书签存在/文本一致/OOXML 基线一致）任一失败即致命；插入点 = 选区 end（非折叠选区不毁文本）；`previousMode` **恢复**（不强制 off）；插入抛错 → **永不重抛**，`partial` + "The image may already be inserted; review before retrying."；插入后 ≤450pt 等比缩放，alt text 尽力而为。SVG 源另存入 customXmlPart（`claric-svg:` alt-title 指针，svg-source-store.js）——PNG 回读会丢矢量源，此存储使后续 `edit_illustration_text` 可行。

### 7.9 IMAGE_TOOL / DOCUMENT_IMAGE_TOOL（图像工具循环，agent-actions.js）

`prepareImageToolEdit`：快照所有内嵌图片（选区或全文档；上限预览 50、全量可寻址，image-model.js:239-240）→ 草稿模型（**永不直接动 Word**）→ ReAct 循环（`tool-loop.js`）。

**工具集**（11 个，agent-actions.js:835-993 分派）：
| 工具 | 执行者 | 特殊门 |
|---|---|---|
| list_images | 草稿模型 | — |
| read_image | **宿主**（Word.run 取 base64）| 观察作为 `image_url` part 附给下一轮；纯文本后端 4xx → **剥离图像重试一次**并标记 `markVisualInputUnavailable`（agent-actions.js:171-195） |
| design_illustration | 宿主嵌套 LLM（SVG）| — |
| replace_illustration | 宿主读源图 + 嵌套 LLM | 图注评审语境（FIGURE_VISUAL_RE）需先成功 read_image；4xx 且无源 SVG → 拒绝不降级（:478-480） |
| edit_illustration_text | 宿主确定性 DOM 编辑 | 无 LLM |
| edit_figure_caption | 草稿模型 | **最强门**：非活索引/未 read_image/视觉不可用/位置非 before-after/距离≠1/before-After 空/超 4000 字/相同 → 拒；候选段必须匹配题注文本正则 + 强度 + 未截断 + OOXML 可用 + 非字段段/非含图段 + 归一化相等；**evidence 与 style 必须逐字复制自候选**——模型必须"证明读过"（image-model.js:329-391） |
| delete/resize/align/set_alt_text/set_image_link | 草稿模型 | resize 三选一（宽/高/比例）；link 仅 http(s)/ftp/mailto/file/# |

循环机制（tool-loop.js）：步预算 **image=8**（table=14，默认 12，agent-actions.js:45）；协议违规/未知工具 → **错误观察**让模型自纠（不抛）；同一调用指纹连续 3 次 → `repeat-limit` 提前收场（:372-389）；`finish.summary` 缺失收 `''`（:360）；预算耗尽 → `step-limit`；**只读收场（无 ops + summary）→ `{noOps:true, answer}` → 直接作为聊天答案，无卡**（conversation.js:1482-1486）。请求预算：文本 260KB / 带图 6MB+260KB，超限先逐出最旧历史回合再逐出 assistant/observation 对，仍超则抛（:184-233）；超长观察序列化成**合法 JSON 信封**而非切片（:92-115）。

**Apply**（`applyImageOps`，agent-actions.js:1141-1222）：两级 staleness 全致命——图片计数 ≠ 快照计数；逐图 `identityKey` 重算不匹配。索引 op 先行、插入最后（不使已解析索引移位）；`attempted` 一次性守卫（模块级 WeakSet）；部分失败语义同 FORMAT（首写后出错 → 警告 + partial）。

### 7.10 TABLE_TOOL / DOCUMENT_TABLE_TOOL（表格工具循环）

`prepareTableToolEdit`（agent-actions.js:237）：从选区表格区域（或 `readDocumentTableRegions` 全文档，每表逐个 sync）种子草稿模型 → 12 个纯函数工具（get_state/set_cell/insert_row/delete_row/merge_cells + 7 个样式工具，table-model.js:54-115）。

**关键校验**（table-model.js）：
- 行 op 授权在**准备时**计算：`!merged && bounds 覆盖全宽`（:133）；合并中的表禁行 op；一次性仅一个合并；`deleteRow` 不允许删光所有行。
- 单元格坐标恒指**原始行号**——行 op 排队不即时移位（:22-26）。
- 样式 op 共享预算门；`setHeaderRow` 仅首行；`setColumnWidths` 拒绝合并表、需恰好 colCount 个宽度。
- `toTablePatch()`：过滤"行将被删除"的 cell 编辑；行 op 转 planRowOpOrder 降序；单表剥 `tableIndex` 保持旧形状，多表全量携带。

循环失败模式与 IMAGE_TOOL 相同（协议观察化、repeat-limit、step-limit）。**unparseable 单发 patch 的重试入口在 SELECTION_EDIT（§7.2 第 3 步）**。翻译出的 patch 复用 tablePatch 形状 → 同一张卡、同一个 `applySelectionAmendment` 表格分支（conversation.js:1625-1645 → `_stageTablePatchProposal`）。

文档级：一个循环跨所有表（模型用 `tableIndex` 选表）；卡项前缀 `TN:` 标明归属表；Apply 侧每个 op 锚定 `body.tables.items[i]`。

### 7.11 COMPOUND（复合指令，conversation.js:1997-2057）

1. `planDocumentTasks`（task-planner.js）：意图级 prompt（**规划器永不看文档文本**，:17-19）；输出契约 = 纯 JSON 数组 `{type, instruction}`；允许 9 类；≤6 任务；单条指令 ≤500 字（超出截断保留）。
2. 校验分支（:105-142）：falsy/非数组/JSON 错 → null；未知 type / 空指令 → 单项丢弃（未知 type 有日志，非对象项静默丢）；0 幸存 → null；>6 → 截断。
3. **null → 回退单意图重路由**：`routeTurn(原文, {...selectionFacts, skills: [], allowCompound: false})`——skills 清空防止再次匹配 /skill；dispatch 到同一消息上（conversation.js:2016-2031）。
4. 成功 → 逐任务 `dispatchTurn(turnForTask(task))`，每个任务**自己的提案卡**（文本修订与图/表 op 永不混卡）；任务间重申 busy；**任一任务被取消 → 跳过剩余全部**（:2046-2049）。
5. 共享 AbortController：cancel 一次停"进行中的子任务 + 所有后续任务"。

### 7.12 SKILL（/skill 分派，conversation.js:1879-1911）

| category × scope | 管线 |
|---|---|
| chat / context（自定义 context 提示词=聊天人格）| runQaTurn（args 并入；无 args 用 skill.description）|
| summary | runSummaryTurn → 新文档（`Application.createDocument`），DOMPurify html profile 清洗，完成报 chars/commentCount |
| tools（保留 /mcp）| runMcpToolsTurn |
| comment | selection-first && 有选区 → **fire-and-forget** 注释（领取后立即释放 busy，多个注释可并行在途；signal 在释放后仍有效——竞态下已 resolve 的请求若 abort 已落，仍不插入，comment-request.js:100-102）；否则文档级注释管线 |
| amendment（default）| selection-first && 有选区 → runSelectionEditTurn；否则 gateApply 文档管线 |

`withArgs`：skill 参数以 `\n\nAdditional instructions from the user: …` 并入模板（:2372-2375）。

**/mcp 分支**（:1804-1877）：无配置 → 指引消息；逐服务器连接，**单台失败仅警告跳过**；全部失败 → "No MCP server could be reached…"；只有资源工具 → "expose no tools"；合成 `mcp_list_resources`/`mcp_read_resource`；工具名合法化 + 命名空间化，**二次碰撞直接跳过而不是误路由**（mcp-tools.js:91）；步预算 = 设置（钳 1-48）否则默认 12；结果上限 64KB 截断、>1.4M base64 图像替换为说明行；**只读契约——MCP 结果只是观察，永不直接写文档**（mcp-client.js:12-14）；URL 白名单（HTTPS/loopback/相对路径）、token 与 URL 全程脱敏。

### 7.13 注释管线（fire-and-forget 细节，comment-request.js:40-160）

选区捕获为书签 → 排队（≥5 条警告 "LLM may slow down"）→ LLM → `restoreProtocolMarkers` **在响应返回后才执行** → 插入书签范围 + 删书签。分支：书签捕获失败 → 移出队列 + 错误日志；响应后 abort → "Comment request cancelled — nothing was inserted."（**无 Retry 链接**——用户主动取消不重试）；插入时书签丢失 → 警告并**记录完整 LLM 响应文本**不丢失；其他错误 → Retry 闭包复用**原书签**（不用当前选区重捕）、**不传原 signal**（早已中止，传了必失败，:145-157）。

### 7.14 summary（新文档）

`runSummarySkill`：extractAllComments（WordApi 1.4 `body.getComments()`，关联回复按创建时间重排——MS 不保证时序；配对数不匹配**抛错**不静默）+ extractDocumentStructured（优先 accept-all OOXML 段落文本——`para.text` 会把删除线内容与插入内容交错）+ extractTrackedChanges（OOXML 解析，任何失败 → `{changes:[]}`，本流程**唯一**整体吞错的读取器）→ `composeSummaryMessages`（{comments} 追加降级；{whole document}/{tracked changes} 条件替换）→ `marked.parse` → `buildSummaryHtml`（DOMPurify html profile，禁 style/表单/媒体标签；`<img src=data:>` 因 DOMPurify 对 DATA_URI_TAGS 的行为被评估为惰性保留，document-generator.js:14-19）→ 新文档打开。

---

## 8. 模型传输层

### 8.1 请求构造（lib/llm-client.js）

- 分派：`apiFormat === 'anthropic'` 或 claude preset → Anthropic 原生 Messages API；其余 OpenAI 兼容 `/chat/completions`（:278-282, :320）。
- URL 拼接防双前缀：base 已以 `/v1` 结尾则不再加（:258-263）。
- Anthropic 头：`anthropic-version: 2023-06-01` + `anthropic-dangerous-direct-browser-access: true` + `x-api-key`（第三个头正是 claude preset 唯一 `staticOk:true` 的原因——providers.js:129）。OpenAI 系：`Authorization: Bearer`，**key 为假则整个头省略**（本地 Ollama 免钥）。
- Anthropic 体（:911-943）：max_tokens 必填（思考预算 → max(16384, budget+8192)；xhigh/max → 65536；否则 16384）；system 提升为顶层；非 assistant 角色折叠为 user；温度**钳到 [0,1]**（1.8 静默变 1.0）。OpenAI 系温度越界则**替换为 1**（:348-350）——两路径策略不对称（§12）。
- **工具调用从不进请求体**——工具由 tool-registry 写进 system prompt，纯文本 ReAct（面向任意 OpenAI 兼容端点）。
- 模型能力（model-capabilities.js）：有序匹配表（provider 且 model regex）决定思考字段映射（15 种协议：reasoning_effort / chat_template_kwargs / thinking.type / output_config.effort …）；`requested === 'default' → 发送空对象（什么都不发）`（:705）——某些后端见到 reasoning 字段即 4xx。温度按 profile/level 抑制（o 系/GPT-5/Kimi K 系不支持；OpenAI 推理系仅 effort=none 时接受）。
- 图像附件：`image_url` part → Anthropic `_anthropicContent` 翻成 base64 块；**system 消息里的图像被静默丢弃**（:889-896）。

### 8.2 流式（:611-806）

- SSE 手工解析：`data:` 行、`[DONE]`、delta 取 `choice?.delta ?? choice?.message`（容忍整消息帧）；`finish_reason` 锁存；reasoning 三种厂商拼写（reasoning_content / reasoning / reasoning_details[].text）。
- **`createStreamDemux`**（:196-243）：增量 `<think>` 分流，tag 跨 token 时按最长后缀 holdback；未闭合 `<think>` 的残文归 reasoning（与 stripThinkTags 一致）。
- **空闲超时而非总时长**：每次收到 chunk 重置计时（:638-645, :742）；聊天 UI 传 300s，即"模型 5 分钟不吐字才超时"。
- Abort：本地 AbortController + 外部 signal 转发（刻意不用 `AbortSignal.any`——WebView2 兼容，:503-505）；已中止的 signal 进入即同步抛。
- 服务器忽略 `stream:true` → 整体 body 当一帧，不报错。
- `body` 不可读（含 Jest mock）→ 非流式回退。

### 8.3 错误分类总表（llm-client 层**零重试**；一切重试都在上层）

| 条件 | 传播形态 | 上层处理 |
|---|---|---|
| Abort（用户）| AbortError 原样 | 各 runner → "Cancelled." |
| 空闲/总超时 | TimeoutError + 具体消息 | markError |
| HTTP 4xx/5xx（同一形态）| `HTTP <status> <text>: <body[0..300]>` | 上层用 `/^HTTP 4\d\d/` 识别 4xx 以决定多模态降级 |
| 网络/CORS/DNS | fetch 原始 TypeError（**无 CORS 专项分支**）| markError；缓解靠架构（origin 自适应代理预设），不靠错误处理 |
| SSE 行 JSON 损坏 | 计数跳过 + 警告；计入截断错误附录 | — |
| 流截断（无 [DONE]/finish_reason）| "LLM stream closed before completion… Retry the request." | markError |
| finish_reason=length / stop_reason=max_tokens | 截断错误 + 建议 | markError |
| Anthropic 流 error 事件 | "Anthropic stream error: …" | markError |
| **空内容** | **成功返回 ''** | 各管线自行检测（append→"no content"；orchestrator 记无修改；QA 显示空） |
| 非流式响应 JSON 损坏 | 原始 SyntaxError | markError |

`_describeHttpError` 先读 body 再格式化（HTTP/2 常无 statusText）；读 body 失败也不吞状态行（:398-416）。

### 8.4 JSON 修复（lib/json-utils.js）

`extractJsonObject`：候选源 = 每个 ```json 围栏 + 裸文本（末位）；逐源直析 → 字符串感知的平衡 `{}`/`[]` 扫描（mismatch 闭合终止候选而非毁栈）→ 必要时**字符串感知**地去尾逗号（朴素 `,\s*}` 会在字符串字面量内删字符——表格单元格文本会被静默破坏，:10-15 注释）。**完整顶层数组不拆内层对象**（完整响应 ≠ 散落对象）。失败抛三种消息（"not an object" / 上次解析错误 / noObjectMessage），可定制。`extractJsonArray` 同型但返回 `{value, error}` 且拒绝从对象字段里捞数组。

### 8.5 图像生成错误面

所有错误经 `sanitizeImageErrorMessage`：剥 API key、请求 URL → `[redacted URL]`、data URL → `[redacted data URL]`、Bearer/query 凭据（:102-121）。**副作用：图像 4xx 不带 `^HTTP 4\d\d` 前缀**（变成 `Image request failed: HTTP 4xx …`）——LLM 多模态降级 regex在此层不适用（该层也确实无需降级）。

### 8.6 testConnection

`GET {base}{apiPath}/models`，30s 固定超时，**对 Anthropic 也用 OpenAI models 端点**（仅 auth 头不同，:1237-1238）。图像连接测试无廉价 ping → 发起**真实微型生成**（灰圆，60s，image-client.js:481-492）。

---

## 9. 提案卡与用户审阅

### 9.1 守卫链（proposal-card.js:479-530，时序严格）

1. 重入/已结算/正在应用 → 拒。
2. 全部取消勾选 → Apply 禁用。
3. **`isBlocked()` 预检**（conversation.js:781-783 注入）：epoch 过期 → "This proposal belongs to a previous chat…"；busy → "A run is currently processing…"。**拒绝先于 controller 注册**——进行中的 run 保住自己的 controller 和 busy 标志（:482-493 注释）。
4. **跨卡互斥** `_anyCardApplyInFlight`：另一张卡在应用 → 拒。
5. 通过 → `setState('applying')` + 新 AbortController 经 `registerController` 注册进 appState（Stop 按钮可暂停进行中的应用；防御：**绝不覆盖外来 controller**——那会让该 run 不可取消，conversation.js:800-803）。

### 9.2 终态语义（"Applied 不许说谎"）

| 终态 | 卡上文案 | 语义 |
|---|---|---|
| markApplied | "Applied {n} of {m} change(s) as tracked changes." | 全部选中项落盘 |
| markRejected | "Rejected — no changes were made." | 经公开 API 走（直接 settle 会把历史卡永久留在 pending，:576-583） |
| markWarning | 自定义 | **部分/零落地**——诚实呈现（表格 patch 诚实计数、chunk apply 的 errors 逐条入日志） |
| markError | "Apply failed: {msg}" | Apply **重新启用可重试**；已部分应用则 Reject 禁用 |
| setPaused | Apply 变 "Continue applying" | Stop 中断；Reject 禁用（"部分应用后拒绝会误导"）|

### 9.3 逐项进度

应用回调 `onChunkApplied` → `markItemApplied(id, {applied|noChange|skipped|error})`：勾选+禁用+置灰+标签（:421-442）。未勾选项在应用前被 `filterTablePatchBySelection` / chunkIds 过滤——**只写勾选的**。

### 9.4 auto-apply（自动应用）

开关需先开 Track Changes（否则 "Auto-apply requires Track Changes…"）+ 模态确认（确认器抛错 → 保持关闭）。回合落定后 `setTimeout` 顺序 drain pending 卡（chat-view.js:861-871, :902-917）；会话切换/消息出错/设置翻转即停。**关闭 Track Changes 强制关闭 auto-apply**（settings-view.js:314-324）。

### 9.5 渲染安全

- 聊天文本一律 `textContent`——**无 markdown 引擎，模型输出不能注入 markup**（chat-view.js:8-9）。摘要在 `llm-client.stripMarkdown`（6 遍去 fence/标题/加粗/斜体/项目符号/反引号，保留编号列表）供 history 用。
- 唯二 innerHTML：sanitizeSvg 后的插图预览（渲染时**再消毒一次**，proposal-card.js:295-304）与图像 diff After 面板。
- diff 用 diff-match-patch DOM 节点（del/ins），仅展示。

---

## 10. 写回 Word：diff、对齐与 tracked changes

### 10.1 diff 策略链（lib/word-diff/）

```
选择策略（word-actions.js:1281-1290）:
  lineDiffEnabled            → sentence-diff（强制句级）
  否则任一侧含 CJK           → char-diff（字符级；词级会把中文一句当一个 token，
                                改一个逗号变成整句红线——char-diff.js:4-9）
  否则                       → token-map（词级）

失败降级链: token-map → sentence-diff → block-replace（删原文+后插新文）
每次降级前：范围重置为原文且 tracking off（重置本身不得成为假修订）
block-replace 也失败 → "All diff strategies failed. Final error: …"
```

关键工程细节：
- **出现序号技巧**：Word search 子串匹配（"the" 命中 "other" 内部），`_occurrenceIndex` 模拟 Word 的贪心非重叠扫描选对第 k 个命中（char-diff.js:76-87）；不可解 → 抛入兜底。
- 删除 token **合并成连续 run**——逐 token 删除会撑爆 Word 会话级 undo 记账导致应用后编辑卡顿（token-map.js:194-214）。
- char-diff `MAX_OPS=200` 上限即抛；搜索片 ≤200 字符**永不切开代理对**（emoji/CJK 扩展）。
- sentence-diff 修正上游 bug：diff 输入必须用**出现序**序列而非去重序列（否则重复句静默错位，:11-15）；退化布局（单一"句"=CJK 无句界 / 全删）→ 直接 block-replace。
- 明确警告禁用 `match.getRange(after)`——搜索产出的范围上它返回零宽点，`.text === ''` 破坏后续定位（char-diff.js:257-263）。

### 10.2 段落对齐（reassembler `_alignParagraphs`，:171-281）

1. LCS（精确修剪文本）锚定未变段落——O(m·n)，受 chunk 大小约束。
2. 锚点间隙内贪心前向匹配，相似度阈值 0.4；单步前瞻决定 insert/delete，否则 delete+insert 对。
3. 相似度：间隔文字用词重叠，**CJK 用字符 bigram Dice**——CJK 无空格，整段一个"词"，词重叠恒 0，所有润色段会被对成 delete+insert（:97-115）。
4. 失败分支：无段落 → 抛（降级范围策略）；全空白 → no-change；**<30% 截断 → TruncatedOutputError（无兜底）**；全等 → no-change；单段 diff 失败 → 整段替换（丢 run 格式保段属性）；对表格段的 delete/insert 跳过警告（删单元格内容 ≠ 删行，删了就毁表）。
5. 性能：变更段的内容范围**一次批量预解析**（原始文本尚在时），逆序编辑互不移位——注释明确记载逐段 round-trip 曾主导整文档应用耗时（:414-431）。

### 10.3 书签生命周期（两套独立系统）

**chunk 书签 `_wdp*`**（staged 文档修订）：
- 命名：`_wdp` + hex 时间戳 + hex 序号 + 3 随机 alnum（下划线前缀=隐藏）。
- 创建在 LLM 处理**之前**（熬过处理时长）；创建前校验两端边界段文本（chunk 索引来自上一次 Word.run，其间任何增删都会静默移位）。
- 应用时 re-anchor：在书签范围内定位存储的段落序列窗口（修剪比较、忽略空白段）——期间用户插入标题等导致范围"吸收"了新段落也能收窄到正确窗口；找不到窗口 → **跳过该 chunk**（不冒险删）。
- 清理：应用成功 → 删（失败 chunk 的**保留**供 retry）；中断 → 全保留供 Continue；拒绝 → discard 全删；**启动时孤儿回收**（taskpane reload 摧毁了 apply/discard 闭包的场景，:615-623）。
- 丢失 → 该 chunk 'bookmark range lost' 跳过，绝不妨碍其他 chunk。

**锚点书签 `_claric_fmt_*` / `_claric_img_*`**（格式/插图）：
- 创建于 LLM 调用前；配 text +（尽量）OOXML 基线；插图 cursor **强制**要 OOXML 否则不建。
- 任何 apply 尝试后即弃（`cleaned` 幂等）；丢失/漂移/基线不符 → **一律致命**（无部分模式）。
- 宿主能力在 prepare 时检查（缺 insertBookmark 等 → "This Word host cannot anchor formatting safely."）。

### 10.4 平台与能力差异（fail-closed 原则）

| 能力 | 检测方式 | 降级行为 |
|---|---|---|
| 行插删作为修订 | `platform.js: supportsTrackedRowOps` = PC\|\|Mac；**unknown 也算不支持**（"静默不追踪的结构编辑比告知更糟"）| 非桌面：行 op 不追踪运行 + 警告"Row insertions/deletions cannot be tracked… Cell text edits are still tracked." |
| 表格插入追踪 | 同上 | 警告 + 直接插入 |
| 单元格合并 | —（任何宿主都不追踪）| 永远直接 + 警告 |
| WordApi 1.4 注释 | `isSetSupported`（启动）| 注释回合报需求；amendment 附带注释 → 记日志不插入 |
| WordApi 1.3 表格/列表 | isSetSupported / `typeof Word.List` | 回合报错 / 警告跳过 |
| `Word.ChangeTrackingMode` | 真值检测（**无一处用 isSetSupported**）| 不设置 tracking；建表时追加专门警告 |
| insertRows / merge / getBorder / autoFitWindow / distributeColumns / setCellPadding / table.font | `typeof` 特性检测 | 逐项警告跳过或回退（如 per-row font 回退） |
| imageFormat（WordApiDesktop 1.1，web 上 sync 即抛）| 独立 sync + 失败不读 | agent-actions.js:402-409 |

两种 tracking 恢复纪律并存：**强制 off**（amendment/表格/append/cleanup——注释与后续回合不得继承）与**恢复 previous**（format/插图/image op——尊重用户原模式）。

---

## 11. 横切机制

### 11.1 取消模型

- 单一 `cancel()`（conversation.js:2305-2316）同时 abort `processDocController` 与 `chatController`；aborted 计数只用于日志。
- 信号检查点遍布：每个 runner 的 `_beginChatTurn`、每 chunk 边界、每个 Word.run 内多处 `checkOperationSignal`（format 有 8 处，让 Stop 落在 op 之间）。
- **中止的语义分层**："已落盘的改动保留"（文档 run 中止 → "Cancelled — already-applied changes remain…"）；注释请求中止 → 不插入；卡 Apply 中止 → setPaused 可续。
- **"取消是用户的决定，永不静默重试"**——图像 fallback 与 comment retry 均显式重抛 AbortError（word-actions.js:3023-3024；comment-request.js:136-142）。

### 11.2 busy/互斥不变量

`isBusy() = submissionOwner || isProcessing || isProcessingDoc || isProcessingSummary || processDocController || chatController`（conversation.js:714-716）。关键不变量（makeProposalCard 注释，:765-814）：**每张卡的 apply 都必须经 registerController 领锁**——曾经只有文档修订卡接线，其他卡的应用在 busy 标志 DOWN 时运行，能与新回合的 parse→bookmark 阶段竞态（或与第二张卡的应用竞态）。

### 11.3 存储与配额（sessions.js）

写入前逐级降级（:132-179）：①去插图预览 → ②diff 项截 2000 字符 → ③整卡丢弃（最旧先）→ ④消息文本截 100K。然后：会话数 >50 → 弹最旧；**总字节逐出先于新 blob 写入**（"先写可能在旧 blob 还占着即将腾出的空间时就把配额炸了"，:300-330）；仍失败 → 删光其他会话再试；索引还失败 → "Session recovery failed…"。blob 成功索引失败 → "Session saved, but the history index could not be updated… History may be stale."

### 11.4 安全

- **模型输出注入面**：聊天纯 textContent；SVG 双重消毒（生成时 + 渲染时）；摘要 HTML DOMPurify html profile（禁 style/表单/iframe…）；`<img data:>` 被记录为惰性例外（document-generator.js:14-19）。
- **协议注入面**：文档文本可伪造 `===AMENDMENT===`/`[END TEXT]` → 一律 defang 进 prompt、先解析后 restore（两处独立注释强调顺序，orchestrator.js:108-118, :312-316）。
- **网络面**：图像/MCP 端点白名单（HTTPS/loopback/相对代理；拒控制字符、反斜杠【WHATWG 会当作 authority 分隔符】、`//` 前缀、内嵌凭据）；错误消息脱敏（key/URL/Bearer/query token）。
- **MCP**：只读观察契约、名字空间防误路由、token 全链路 redact、-32601 优雅降级。

---

## 12. 发现的边界情况与不一致

以下为审读中确认的事实（非推测），按影响排序：

1. **非 Word 宿主 / Office.js 缺失 → 面板静默死亡**（taskpane.js:36-42）：无任何用户可见提示，只有静态未接线的 welcome 标记。
2. **`mode === 'both'` 不可达但代码仍分支**：`getActiveMode` 因 amendment/comment 互斥永返不了 'both'（prompt-manager.js:244 注释自认），而 orchestrator.js:100/:309 仍保留 `mode === 'both'` 分支——活死代码，merged 实际只由 `amendment + commentInstructions` 路径触发。
3. **`{whole document}` / `{tracked changes}` 静默字面量**：调用方漏传数据时占位符原样保留，无追加降级、无警告（prompt-manager.js:475-482）；且 summary 路径**不做 defang**（只有 composeMessages/composeMergedMessages 做）——插值进去的文档文本理论上可携带未中和的协议标记。
4. **工具循环 `finish` 无 summary 被接受为 `''`**（tool-loop.js:360），尽管 prompt 声明必填；两个调用方随后把"只读 + 空 summary"当作无改动抛 `noChanges`（agent-actions.js:295, :1080）——模型不守协议时用户看到的是报错性文案而非答案。
5. **Sanitizer 清空 SVG 无专属错误**：`sanitizeSvg` 返回空/非 `<svg>` 时坍缩为通用 "The design step produced no usable SVG."（§7.8），可诊断性缺口（安全无虞——不会插入任何东西）。
6. **Anthropic 与 OpenAI 的越界温度处理不对称**：Anthropic 钳制（1.8→1.0，静默），OpenAI 替换为 1（不同静默语义）；未见测试或注释证明是有意的（llm-client.js:940 vs :348-350）。
7. **`sendPromptStream` 丢弃 reasoning 通道**（llm-client.js:807-816）——QA/append 的思考链对用户不可见（与 sendMessagesStream 路径不一致）。
8. **附件上下文整块截断**：`buildAttachmentContext` 在首个超限块处**整体 break**——一个超大文件会关掉其后所有附件的上下文（file-attachments.js:165-173），而非逐块均匀分配。
9. **`settings-loader` 无提交时校验**：后端未配置时提交照跑，错误延迟到回合内才暴露（§3.5）。
10. **UI 无任何 token 计数**："Model activity · N sections"、"Worked for Ns · N steps"——只有 settings 里的历史预算输入有 token 字样（chat-view.js:717-735）。
11. **`getComments()` 文档顺序是经验事实非契约**：comment-extractor.js:369-375 自认靠 Office.js 未文档化的行为，下游 splicer 以运行时不变量检查兜底。
12. **history 回放不含图片**：历史条目强制 string content（conversation-history.js:69-72）——多轮对话中用户上一轮发的图在下一轮请求里只剩名字注记。设计使然（预算），但值得知晓。
13. **`_alignParagraphs` 用 `alignment.indexOf(op)` 定位插入锚**（reassembler.js:503）——依赖 op 对象恒等（从不克隆成立），O(n) 每插入；脆弱耦合但当前正确。
14. **无任何 llm-client 层重试**：截断/超时全部直通 markError，只给"Retry the request"文案由用户手动重发（多模态 4xx 与 chunk retry 除外）。

---

## 13. 附录：关键常量表

| 常量 | 值 | 位置 |
|---|---|---|
| 选区图片快照上限 | 6 | word-actions.js:839 |
| 选区变化 debounce | 200ms | word-actions.js:971 |
| 附件 | 5 文件 / 10MiB 文本 / 4.5MiB 图 / 10MiB 总 / 200K 字符 | file-attachments.js:31-37 |
| 文档 chunk | 6000 tokens；并发 6（任一 chunk>8000 → 4）；每 chunk 300s | word-actions.js:3277-3300 |
| QA/聊天超时 | 300s 空闲 | word-actions.js:3743 |
| 客户端默认超时 | 非流式 120s；流式=空闲 | llm-client.js:437/:611 |
| 会话历史预算 | 默认 128000 tokens − 8192 保留，钳 1000…2M | conversation-history.js:58-67 |
| 历史截断 | 字段 1000 字符；提案 8 / 项 8 / 附件 8 | conversation-history.js:27-30 |
| 规划器 | ≤6 任务；指令 ≤500 字符；9 种类型 | task-planner.js:29-38 |
| 工具循环步预算 | table 14 / image 8 / 默认 12 / MCP ≤48 | agent-actions.js:45; tool-registry.js:21; conversation.js:1813 |
| 工具循环请求预算 | 响应 64KB；观察 64KB；带图 6MB+260KB | tool-registry.js:21-50 |
| 重复调用熔断 | 连续 3 次同一调用 | tool-loop.js:375-389 |
| 插图 | SVG ≤256KB；源 SVG ≤50KB；插入宽 ≤450pt；图像 8M base64 字符 | illustration.js:23/:56; word-actions.js:3059; image-client.js |
| 图像生成超时 | 180s（连接测试 60s） | image-client.js:389/:481 |
| MCP | ≤10 服务器；结果 64KB；图像附件 1.4M；请求 30s | app-state.js:242-252; mcp-tools.js; mcp-client.js |
| 会话存储 | 50 个 / 1.5MB 每个 / 4MB 总量 | sessions.js:27-29 |
| 活动日志 | 200 条 | status-bar.js:13 |
| 插图位置锚 | 光标走查 ≤120 段找最近标题 | word-actions.js:706 |
| 对齐相似度阈值 | 0.4（词重叠 / CJK bigram Dice） | reassembler.js:107-115 |

---

*报告完。所有 `file:line` 均指 main 分支（693e447）源码；四条并行审读线索（模型传输层 / Word I/O 与应用层 / 编排与工具循环 / UI 与状态层）+ 路由与管线主干的人工直读交叉验证。*
