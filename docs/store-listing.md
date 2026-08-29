# Microsoft Marketplace — Store Listing Draft (Claric 1.0.0)

Submission copy for Partner Center → Marketplace offers → Office Add-in.
The tone rule for everything below: **plain language, written from the
user's pain point, zero jargon-first framing.** Feature lists are framed as
"what you get to do", not as architecture.

---

## Store name

**Claric — AI Redlining for Word**

> Naming rationale: the brand stays short and pronounceable (from
> *clarify*); discoverability comes from the "AI Redlining for Word"
> descriptor, which is also the compliant pattern — store policy forbids
> branding the product as "Word"/"Office" itself, but "X for Word"
> suffixes are accepted. The name is env-overridable via `DISPLAY_NAME`
> (see `.env.example`), so a rename is a one-line change.

## Short description (≤ 250 chars, shown in search results)

> An AI editor that lives in Word. Ask in plain language, review every
> proposed change as native tracked edits, and apply only what you
> approve. Works with the AI you choose — local models or cloud.

## Full description (EN — primary store language)

**Your document. Your call. Nothing changes without your approval.**

You know the scene: a 40-page contract due tomorrow. Vague clauses,
inconsistent numbers, formatting chaos from five rounds of edits — and
every fix must be traceable, because someone senior is going to read the
redlines.

Claric puts an AI editor **inside Word** that works the way careful people
actually work: **you ask in plain language (English or 中文), it proposes,
you decide.**

### What you can ask it to do

- 🖊️ **"Polish this paragraph."** Select any text, get a cleaner version —
  shown as a before/after diff, applied as word-level **tracked changes**
  your colleagues can accept or reject, just like a human editor's redlines.
- 🔍 **"Check this whole contract for risky terms."** Claric reads the
  document section by section, leaves margin comments where it sees
  problems, and proposes fixes — each one a separate review card.
- 💬 **"Summarize all the comments in this document."** Every reviewer note
  distilled into one clean summary you can act on (or forward).
- 📊 **"Add a 4×3 table with our project milestones."** Tables appear as a
  preview first — no dragging grid handles, no misaligned cells.
- 🎨 **"Design an illustration for the cover." / "Center all the
  headings." / "Continue writing from here."** Formatting, figures, and
  first drafts — all on request.
- 🧹 **"Delete the empty paragraphs."** The tedious cleanup that used to
  take twenty minutes of squinting.

### Nothing is applied until YOU approve it

Every AI suggestion lands as a **proposal card** in a chat panel beside
your document: the before/after diff, one checkbox per change, a button to
jump to the exact spot in the document. Tick what you want, apply — Word
records it all as **native tracked changes**. Reject the rest with one
click. Your document is never silently rewritten, and your co-workers see
a normal, reviewable redline.

### Your AI. Your rules.

Claric doesn't lock you into one AI vendor — you connect the model:

- 🔒 **Run AI locally** (Ollama, vLLM) so sensitive contracts never leave
  your computer.
- ☁️ **Or use a cloud model** — DeepSeek, Zhipu GLM, Moonshot Kimi, MiniMax,
  or any OpenAI-compatible API — with your own key. You're in control of
  what's sent and where.

In legal and compliance work, an edit you can't see is an edit you can't
trust. Claric keeps every AI suggestion visible, reviewable, and yours to
accept.

**Try it now:** open any document, select a paragraph, and type *"polish
this selection."*

## 中文参考版（如提交多语言 listing）

**你的文档，你说了算——AI 不经你确认，绝不动一个字。**

想象一下：明天要交一份 40 页的合同，条款含糊、数字前后不一、改了五轮的
格式一团乱——而且每一处修改都必须留痕，因为合伙人要看修订记录。

Claric 是住在 Word 里的 AI 编辑，按严谨的人的方式工作：**你用大白话下指
令（中文英文都行），它出方案，你拍板。**

- 🖊️ **"润色这段话"**——选中任意文字，给出更通顺的版本，以逐词修订
  （tracked changes）方式落地，同事看到的和人类编辑的红笔批注一样。
- 🔍 **"检查整份合同的风险条款"**——分节通读全文，在风险处加批注，
  并给出修改建议，每条建议单独一张确认卡。
- 💬 **"总结文档里所有批注"**——几十条评审意见，一分钟变成一份可执行的摘要。
- 📊 **"插入一张 4×3 的项目里程碑表格"**——先看预览再插入，不再手动拖格。
- 🎨 **"给封面画一张插图" / "所有标题居中" / "从这里继续写"**——排版、配图、续写，一句话的事。
- 🧹 **"删除多余空段"**——过去二十分钟的眯眼排查，现在一键完成。

**所有修改先审后用**：每条 AI 建议都是一张提案卡——改前改后对比、逐条勾
选、一键定位到原文；应用后全部是 Word 原生修订，可逐条接受或拒绝。文档永
远不会被悄悄改写。

**AI 你说了算（BYOK）**：本地模型（Ollama / vLLM）让敏感合同不出你的电脑；
或填入自己的 key 使用 DeepSeek、智谱 GLM、Kimi、MiniMax 等云端模型。

现在就试：打开任意文档，选中一段话，输入"润色这段话"。

---

## Listing metadata

| Field | Value |
|---|---|
| Categories | Artificial Intelligence (primary) · Productivity (secondary) |
| Keywords | AI writing, contract review, redlining, tracked changes, proofreading, legal tech, document editing, DeepSeek, GLM, Kimi, Ollama, local AI, BYOK |
| Support URL | `https://github.com/Scilence2022/Claric/issues` (set via `SUPPORT_URL`) |
| Privacy policy URL | `https://scilence2022.github.io/Claric/privacy.html` (docs/privacy.html, published via GitHub Pages) |
| EULA | Standard Microsoft contractual terms, or attach the MIT license + disclaimer from `LICENSE` |
| Manifest | Generated with `HOST=<public-domain> PROTOCOL=https HOST_PORT=443 SUPPORT_URL=… APP_DOMAINS=…` — must pass `office-addin-manifest validate` (passing as of 1.0.0) |

## Reviewer notes (paste into Partner Center "Notes for certification")

> This add-in is **bring-your-own-key (BYOK)**: it connects to the model the
> user configures. To review core functionality, open Settings → choose
> **DeepSeek** (or any OpenAI-compatible provider) → paste an API key → then:
>
> 1. Select a paragraph, type "polish this selection" → a proposal card with
>    a before/after diff appears; Apply inserts **native tracked changes**.
> 2. Type "summarize this document's comments" to exercise whole-document
>    extraction.
> 3. Type "insert a 3×3 table" to exercise table creation (deterministic,
>    no LLM needed).
>
> On Word for the web, structural operations (table insert/row ops)
> intentionally apply **untracked with an explicit warning** — the host does
> not support structural revisions. All text edits remain tracked.
> Privacy disclosures: document text is sent only to the user-configured LLM
> endpoint; API keys are stored in browser localStorage of the add-in origin.
> See the privacy policy linked in the listing.

## Screenshot shot-list (≥ 1 required, 1280×800+)

1. Welcome state with skill chips (`/copy-edit`, `/check-doc`, …) over a document.
2. Selection edit → proposal card with inline diff + per-change checkboxes. *(hero shot)*
3. Whole-document proposal card with per-section citation pills.
4. Table creation preview card.
5. Model activity panel streaming reasoning/output.
6. Settings slide-over showing provider picker + BYOK key field.

## Submission checklist

- [ ] `package.json` version = `1.0.0`, manifest `<Version>1.0.0.0</Version>` (done — validator passes)
- [ ] Public HTTPS hosting with a CA-signed cert on a real domain (not localhost)
- [ ] `SUPPORT_URL` set to the GitHub issues page and reachable anonymously
- [ ] Privacy policy published at the listing URL (`docs/privacy.html` via GitHub Pages)
- [ ] Icons present (16/32/64/80/128 — generated, content-hash cache-busted)
- [ ] Screenshots produced per shot-list
- [ ] Reviewer notes + a working BYOK demo key prepared
- [ ] EULA decision (Microsoft standard terms vs. own)
- [ ] `office-addin-manifest validate manifest.xml` re-run against the production HOST build
