# Project: Claric — your redlining scribe for Word

AI-powered Microsoft Word add-in for tracked changes, document summarization, and comment analysis.

## Development Rules

- **Python packages**: Always use a virtual environment (`python3 -m venv .venv && source .venv/bin/activate`) before installing packages with pip. Never install directly into the system Python environment.
- **Build**: `npm run build` or `npm start` for dev server
- **Test**: `npm test` (Jest); targeted image coverage: `npm test -- --runInBand tests/illustration.spec.js tests/image-client.spec.js tests/image-providers.spec.js`
- **Coverage**: `npm run coverage` followed by `npm run check-coverage`
- **Lint**: `npm run lint` (ESLint 9 flat config)
- **Typecheck**: `npm run typecheck`
- **Verify source and build**: `npm run verify` (lint + test + coverage + check-coverage + typecheck + build + verify-build). CI additionally audits dependencies and scans the container; real Word acceptance is separate.
- **Node**: >=22 (see .nvmrc / package.json engines)

## Architecture

- Office.js Word add-in (Claric, runs in WebView2)
- Chat-driven taskpane: `src/taskpane/` split into `app-state.js`, `skills.js`, `conversation.js`, `word-actions.js`, and `ui/*`; `taskpane.js` is bootstrap only
- Chat providers (11): Ollama, vLLM, OpenAI, Claude (Anthropic), DeepSeek, Zhipu GLM, Moonshot Kimi, MiniMax (international), MiniMax China, 中科大模型 (zhongkeyu.com), and Custom (OpenAI-compatible)
- Provider adapters keep the shared chat contract while handling provider-specific API formats and model controls; Claude uses the native Anthropic Messages API
- Image generation has an independent provider/model route: OpenAI Images, Zhipu CogView, MiniMax Images, 中科云 Images (zhongkeyu.com), OpenRouter Images, SiliconFlow Images, and Custom (OpenAI-compatible), with separate endpoint, API key, model, and size settings. OpenAI Images, 中科云 Images, OpenRouter Images, SiliconFlow Images, and Custom use `/images/generations`; CogView uses that endpoint and returns a hosted URL; MiniMax uses `/image_generation` with `aspect_ratio` and base64 output. The client normalizes all response forms to raw raster base64
- Ordinary schematic/illustration requests use a configured image model and insert its raster output; explicit SVG/vector requests use the sanitized chat-LLM SVG route, and disabled or failed image generation falls back to SVG
- Cloud chat/image defaults are origin-adaptive: static hosts use direct provider origins where CORS permits, while the local server uses same-origin proxy paths (`/openai`, `/glm`, `/minimax` for image presets); OpenAI needs `/openai` or another relay because it sends no browser CORS headers, and Custom needs an explicit endpoint or `CUSTOM_PROXY_PATH`/`CUSTOM_PROXY_TARGET`
- Webpack dev server registers the configured provider proxy paths, including the image routes; Custom is opt-in through `CUSTOM_PROXY_PATH`/`CUSTOM_PROXY_TARGET`
- Proxy timeout: 5 minutes (300s)
