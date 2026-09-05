# Dev Harness 本地协议与安全基线

## 交付范围

这是旧 `scripts/dev-e2e-middlewares.cjs` 的受保护 HTTP 协议基线，不是新的产品后端，也不是 Word 自动化 driver。启用条件仍完全由 webpack 的 `ENABLE_DEV_ENDPOINTS=true` 管理；默认关闭行为未改动。本模块只在 setup 被调用时注册接口。不要将它部署到公网、生产环境或共享工作站服务。

本次没有启动实际用户 harness，没有调用模型服务、Office.js 或真实 Word。`scripts/dev-harness-fixture.cjs` 是非空、合成脱敏纯文本，只验证 HTTP 请求、状态机及存储往返；`passed` 也是 driver 的声明，服务端不会据此证明文档正确。真实 Word 的选区、格式、修订、撤销、API 成功与最终文本均尚未验收。

## 旧实现问题与新边界

旧实现所有路由使用 wildcard CORS 且无 token；任意可访问端点的客户端都能读写。`global` 保存日志及循环状态；多次 setup 增加进程信号监听器，退出时还会调用 `process.exit(0)`。日志无界累积，每十条全量覆盖文件，部分失败仅打印错误却仍返回成功。

现在每个 Express app 只安装一次，运行状态属于该实例；不读写 `global.e2eLogs` 或 `global.e2eLoopControl`，不增加/删除任何进程监听器，不接管退出。所有修改立即提交 snapshot，无需依赖退出补刷。同一存储目录仅支持一个服务实例/进程写入，不提供跨进程锁、并发外部文件编辑或集群一致性。重复 setup 返回原 middleware 数组，不重置状态、不重复显示 token；新 app/restart 则明确重新暂停。

## 授权与 Origin

- 每个非 OPTIONS 请求，包括 GET/DELETE，必须带 `x-claric-harness-token`。缺失或不匹配返回 `401`，不读写业务文件。使用定时安全字节比较，token 不接受 query string、cookie 或其他 header 替代。
- 启动可以传入进程环境变量 `CLARIC_HARNESS_TOKEN`，要求 32–256 个非空格可打印 ASCII 字符。未提供时通过 `crypto.randomUUID()` 创建。不要将 token 写进仓库、文档或 fixture。
- 成功 setup 时在控制台显示 token 一次，供本地 driver 使用；没有 token 文件/API，不在后续请求日志中打印 token 或请求正文。终端输出可能被宿主收集，启动日志应视为秘密。重新生成 token 后旧客户端必须更新。
- 有 Origin 时必须精确等于实际请求服务 origin，即 socket TLS 状态推导的协议加 Host，或者属于 `CLARIC_HARNESS_ORIGINS` 的逗号分隔额外列表。配置只接受精确 HTTP(S) origin，拒绝 `*`、`null`、路径、尾部 `/`。同源匹配不使用 `X-Forwarded-Host/Proto`，不因为 `trust proxy` 放宽。
- 无 Origin 的本地命令行 driver 可以访问，但仍必须提供 token。Origin 绝不是身份认证，也不能阻止能伪造 Host/Origin 的非浏览器客户端；Host 校验仍依赖 webpack 的 allowedHosts 边界。本接口不支持反向代理透明转发，额外 origin 必须显式配置。
- 所有已知路径共用 CORS，包括错误响应和 OPTIONS。覆盖上游可能留下的 wildcard/credentials CORS；允许的 Origin 被精确返回并添加 `Vary: Origin`、`Cache-Control: no-store`。拒绝 Origin 时不返回 Allow-Origin。
- OPTIONS 不要求 token，返回 `204`，不读写业务数据；仅允许 GET/POST/DELETE/OPTIONS 和 Content-Type、x-claric-harness-token 预检 header。失败预检为 `403`。真正请求仍需 token。

## Body 与错误

有 body 的请求必须是 `application/json`、顶层对象，不能是数组、null、字符串或数字；解析错误、非对象或错误 Content-Type 返回 `400`。无 body 视为 `{}`，以兼容 trigger/pause/clear/DELETE；有必填字段的业务路由继续验证。上限是 **64 KiB 原始 JSON**，超限返回 `413`；不支持压缩 body，Express 返回 `415`。服务必须在其他 body parser 消费这些路径之前安装本模块。

常规错误体为 `{ success: false, error: string }`；解析、业务校验和持久化错误额外包含 `persisted: false`。未知业务方法为 `405`，未找到用户 prompt 为 `404`。文件读取损坏、权限/空间不足、rename 失败返回 `500`，不会暴露绝对路径，也不会用空数组覆盖损坏文件。容量不足可能返回 `413`。前端不能只检查网络请求是否 resolved，必须检查 HTTP 状态和 success。

## 接口

所有路径保持旧名称。成功写入通常返回 `200`，保留旧响应主字段并增加 `persisted: true`。

| 方法/路径 | 请求与成功响应 |
| --- | --- |
| POST `/log` | 任意 JSON 对象；服务端覆盖 `receivedAt` 和 `seq`。旧空响应改为 `{success:true,persisted:true,seq}`。 |
| GET `/logs?since=ISO时间` | 返回日志数组；无 since 返回有界全部，最多 1000 条。有 since 按 timestamp（缺失则 receivedAt）筛选；非法 since 为 400。增量读取改用 after。 |
| GET `/logs?after=序号` | 返回 entries/nextCursor/oldestCursor/gap；详见日志游标定义。 |
| POST `/logs/clear` | `{}` 或空 body；清空 entries 但保留最高序号，返回 success/message/persisted。 |
| POST `/api/fix-log` | JSON 对象；追加 receivedAt，返回 success/message/totalFixes/persisted。totalFixes 是保留数量，不是历史总数。 |
| POST `/api/trace-log` | `testRunNumber` 为至多 15 位非负安全整数或数字字符串；`trace` 必须是数组。返回 success/message/traceLength/persisted。不再为每个 run 创建文件。 |
| GET `/api/test-cases` | 静态与动态记录合并为数组；不隐式注入 fixture。没有记录时返回空数组。 |
| POST `/api/test-cases` | `original`、`modified` 必须为非空字符串；可选非空字符串 id。生成 expected=modified、createdAt，返回 success/testCase/persisted。 |
| GET `/api/prompts` | 默认和用户数组按 id 合并，用户覆盖默认。 |
| POST `/api/prompts` | id/name/template 必须为非空字符串；按 id 新增或更新用户记录，返回 success/prompt/persisted。 |
| DELETE `/api/prompts/:id` | 只删除用户记录；对应默认项可能重新显示，不能删除默认文件记录。 |

## 日志游标

旧 GET `/logs` 和 `?since=...` 保持数组响应。`since` 只适合兼容性查询：客户端时间戳缺失、无效、重复或时钟偏差都可能影响筛选，不能用它保证增量完整性。增量消费者应改用 `GET /logs?after=0`，返回 `{entries,nextCursor,oldestCursor,gap}`。

- `seq` 是服务端分配的正安全整数，POST `/log` 覆盖客户端提交的同名字段，并在成功响应中返回 seq。缺失 timestamp 或相同 timestamp 不影响排序和读取。
- `after` 为最后已消费 seq（排他）；初次使用 0。每次完整处理 entries 后保存 nextCursor，再用于下一次请求。不能同时提供 after 和 since；非法/重复游标为 400。
- `nextCursor` 是当前提交的最高序号，即使 entries 为空也不会因 clear 回退。`oldestCursor` 是最早保留记录的 seq；没有记录时是 nextCursor+1。若 after 小于 oldestCursor-1，则 `gap:true`，表示已无法完整补齐被 retention/clear 删除的区间。after=0 也不豁免 gap 检测。
- clear 删除记录，但保留最高序号；新日志继续递增。正常 restart 从同一个原子 snapshot 恢复序号及日志，不因进程重启重新编号。读取旧数组时按文件原始顺序分配 1..N，忽略旧记录自带 seq，首次写入提交新格式。迁移前不存在可靠序号，因此不能把旧客户端时间戳转成游标。
- `logs/e2e-test-logs.json` 的磁盘格式从数组迁移为 `{entries,lastSequence}`，最高序号与记录同文件原子提交，避免双文件提交歧义。直接读取该文件的旧工具也必须迁移。HTTP 不带 after 仍是数组。
- after 高于已存最高序号返回 409，可能意味着错误服务、手工重置或存储回滚，需要操作者确认后重新从 0 读取。序号只在同一个未重置存储历史内有意义，不具备跨备份恢复/删库 epoch 身份；删除/回滚整个 snapshot 后不能保证识别所有旧游标。
- 游标不是持久化 exactly-once。无 fsync 保证，客户端断线时 POST 可能已提交，重试可能产生新 seq 的重复业务日志；消费者保存游标和处理副作用也不是一项事务。gap=true 必须作为数据缺口处理，不能声称没有遗漏。

## 受控开发客户端

`node scripts/dev-harness-client.cjs` 是单次 HTTP 协议客户端，不是 Word driver。没有后台轮询、自动 trigger、自动重试、服务启动、Word 启动/注入或自动修代码。只有明确输入 trigger 才创建运行许可；claim/complete 必须提供显式 runId。

先由操作者提供进程环境变量 `CLARIC_HARNESS_TOKEN`（与已启动 harness 一致），客户端不读取 .env 文件，不接受 token 命令行参数，不打印 token；响应中意外回显的 token 也会脱敏。不要把真实 token 放入 shell 命令历史。

```sh
node scripts/dev-harness-client.cjs status
node scripts/dev-harness-client.cjs trigger
node scripts/dev-harness-client.cjs claim --run-id RUN_ID
node scripts/dev-harness-client.cjs complete --run-id RUN_ID --outcome passed
node scripts/dev-harness-client.cjs pause --run-id RUN_ID
node scripts/dev-harness-client.cjs get-logs --after 0
node scripts/dev-harness-client.cjs status --base-url http://127.0.0.1:3000
```

默认 base URL 为 `https://localhost:3000`。仅接受 localhost、127.0.0.1、[::1] 的 HTTP(S) origin，拒绝外部 HTTP **和 HTTPS**、凭据、路径、query/fragment，localhost 网络连接固定 loopback，不经代理、不跟随重定向。TLS 保持证书校验，可通过 Node 正式受信 CA 配置支持开发证书，不提供 insecure 开关。请求有 10 秒总期限、5 MiB 响应上限。成功输出一行 JSON，退出码 0；参数、鉴权、HTTP 非 2xx、连接/TLS/超时、响应解析失败退出码 1。不自动重试状态修改，避免误重复执行。

pause 不带 runId 时仍是显式全局暂停命令；get-logs 不带 after 返回兼容数组。客户端不会持久化游标，操作者负责保存 nextCursor 和处理 gap；passed/failed 只是协议结果声明。

## 一次性运行协议

旧 `canProceed` 只是提示，不能作为执行许可。**driver 必须先成功 claim，再开始任何测试操作。** 同一个 token 持有人属于同一信任域；runId 是一次性协调标识，不是另一种身份凭证。

1. `GET /api/e2e-loop/status` 启动返回 `canProceed:false, waitingForTrigger:true, state:"paused", runId:null, lastIteration:null`。
2. `POST /api/e2e-loop/trigger {}` 创建 UUID runId，状态 pending、canProceed=true；响应包含 runId 和 ISO expiresAt。期限固定为触发后五分钟，包含领取和执行时间，不因 claim 延长。已有 pending/claimed run 时返回 `409`，避免覆盖正在运行的任务。
3. `POST /api/e2e-loop/claim {"runId":"上一步返回值"}` 仅 pending 状态能成功一次。提交 claimed snapshot 后响应 canProceed=false。并发 claim 只有一个 `200`，其余 `409`；必须等待本次 HTTP 成功才执行。重复请求不幂等。
4. driver 完成后 `POST /api/e2e-loop/complete {"runId":"相同值","outcome":"passed"}`，outcome 仅支持 passed/failed。只有当前 claimed 且未超时的 run 可提交；完成后返回 completed、lastIteration（runId/outcome/completedAt）。driver 自报结果不构成 Word 验收证明。
5. `POST /api/e2e-loop/pause {}` 保留旧全局暂停命令语义，停止当前 run 的协议许可；可选 runId 做当前运行检查，推荐 driver 始终携带它。暂停不可能中止已经开始的 Word 操作，driver 必须自行取消/回收。暂停或完成后可再次 trigger 新 run。

runId 缺失/格式错误为 `400`；未知、被新 run 替代、已领取、已暂停、已完成或尚未领取的状态冲突为 `409`；当前 run 到达期限后 claim/complete/带 runId 的 pause 返回 `410`。期限在每次循环接口访问时检查并立即落盘，不安装定时器；没有后台中止能力。旧 run 被新 run 替代后只返回 409，不保存无限历史。循环 snapshot 写入失败为 500，内存状态不前进。

服务重启总是提交 paused/null 的运行边界，不从上次 claimed 恢复。收到请求但响应丢失时客户端无法知道 claim/complete 是否已提交，应查 status 人工或由 driver 协调；不能盲目重试后重复执行。此协议不保证实际 Word 操作的 exactly-once。

## 持久化与保留

- 普通日志仍在 `logs/e2e-test-logs.json`，fix 日志仍在 `logs/fix-logs.json`。trace 改为 `logs/dev-harness/traces.json` 有界数组；运行状态在 `logs/dev-harness/loop.json`。
- 动态测试及用户提示新写入分别为 `logs/dev-harness/test-cases.json`、`logs/dev-harness/user-prompts.json`。新文件不存在时读取旧根目录 `e2e-test-cases-dynamic.json`、`user-prompts.json`；首次写入将数据迁移到新 snapshot，旧文件不修改，随后以新文件为准。默认 `prompts.json`、静态 `e2e-test-cases.json` 只读。
- 每个日志/trace 集合最多 **1000 条且 JSON 最大 4 MiB**，超限淘汰最旧记录；不是永久审计日志。测试/提示自定义集合最多 1000 条、4 MiB，满后拒绝新增，不静默丢弃。读取已有大于 4 MiB 或损坏的 snapshot 会失败关闭，需要操作人先备份并整理，不自动删除旧数据。
- 每次写入序列化有界 snapshot，创建同目录 UUID 临时文件（0600、排他创建），写完后原子 rename 替换；正常读者看到旧版或新版完整 JSON，而不是半写文件。失败返回 500，尽力清理临时文件；进程崩溃或清理失败可能留下 `.tmp`，需要离线维护。
- `persisted:true` 的含义仅是同步文件写入及 rename 已成功返回。**未调用 fsync，不保证断电零丢失，也不保证崩溃时临时文件清理。** 不将这一语义描述为数据库事务或耐久交付。
- 存储拒绝越界路径与现有 symlink 路径；不抵御具有本地文件写权限的攻击者在检查与写入之间置换目录。rootDir 第三个参数只在 NODE_ENV=test 可注入，用于仓库内临时测试目录，不是生产配置项。新持久化 IO 均位于仓库子目录；不写 token。

## 迁移与验证

旧客户端必须补齐 token header/预检配置，并从“轮询 canProceed 后直接执行”改成 runId claim/complete。直接访问 global 的客户端不再支持。消费旧 `trace-log-N.json` 的工具改读 traces snapshot；依赖根目录用户数据新写入的工具也需迁移。不能假设保存全部历史日志、无期限运行、重复 trigger 总成功或 `/log` 始终空响应。

针对性验证命令：

```sh
npm test -- --runInBand tests/dev-harness.spec.js tests/dev-harness-client.spec.js
npx eslint scripts/dev-e2e-middlewares.cjs scripts/dev-harness-*.cjs tests/dev-harness*.spec.js
```

测试使用项目现有 Express、node:http 和 Node 内置 HTTP 客户端，随机本地端口、仓库内临时目录，完成后关服清理；未安装 supertest。覆盖每条路由鉴权/预检、Origin 拒绝、非法及大 body、即时落盘及数量/字节上限、写入与 rename 故障、损坏文件、symlink 防护、进程监听器不变、幂等 setup、一次性并发 claim、超时/过期/重启边界、旧数据迁移和纯文本 fixture 往返。游标测试覆盖缺失/重复/无效时间戳、服务端 seq 覆盖、retention/clear gap、restart 续号、失败提交不前进；CLI 以实际子进程测试全部命令、非零退出、token 脱敏及拒绝重定向。这是协议测试证据，不应转述为真实 Word E2E 或模型效果评测通过。
