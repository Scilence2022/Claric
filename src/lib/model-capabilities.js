/**
 * Model-specific thinking capabilities and request protocol mappings.
 *
 * OpenAI-compatible endpoints share the chat-completion envelope, but their
 * thinking controls are not interchangeable. This module keeps the UI's
 * model-aware choices separate from transport code so every visible option has
 * one well-defined meaning on the wire.
 *
 * Official references:
 * - https://docs.ollama.com/capabilities/thinking
 * - https://docs.ollama.com/api/openai-compatibility
 * - https://docs.vllm.ai/en/latest/features/reasoning_outputs.html
 * - https://api-docs.deepseek.com/guides/thinking_mode/
 * - https://docs.z.ai/guides/overview/concept-param
 * - https://platform.moonshot.cn/docs/api/chat
 * - https://platform.minimaxi.com/docs/api-reference/text-openai-api
 * - https://platform.openai.com/docs/api-reference/chat/create
 * - https://platform.claude.com/docs/en/build-with-claude/effort
 * - https://docs.anthropic.com/en/docs/build-with-claude/extended-thinking
 *
 * @module model-capabilities
 */

/** Values accepted by the persisted provider-level thinkingLevel field. */
export const THINKING_LEVEL_VALUES = Object.freeze([
    'default',
    'off',
    'on',
    'adaptive',
    'always',
    'minimal',
    'low',
    'medium',
    'high',
    'xhigh',
    'max',
    'none',
]);

const LEVEL_VALUES = new Set(THINKING_LEVEL_VALUES);

function makeOption(value, label, description) {
    return Object.freeze({ value, label, ...(description ? { description } : {}) });
}

function freezeOptions(options) {
    return Object.freeze(options);
}

const GENERIC_OPTIONS = freezeOptions([
    makeOption('default', 'Default'),
    makeOption('low', 'Low'),
    makeOption('medium', 'Medium'),
    makeOption('high', 'High'),
]);

const PROFILE_OPTIONS = {
    ollamaGptOss: freezeOptions([
        makeOption('default', 'Default'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
    ]),
    ollamaQwen: freezeOptions([
        makeOption('default', 'Default'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
    ]),
    qwen3Vllm: freezeOptions([
        makeOption('default', 'Default'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low (4K tokens)'),
        makeOption('medium', 'Medium (8K tokens)'),
        makeOption('high', 'High (16K tokens)'),
    ]),
    deepseekV4: freezeOptions([
        makeOption('default', 'Default (High)'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low'),
        makeOption('high', 'High'),
        makeOption('max', 'Max'),
    ]),
    legacyDefault: freezeOptions([
        makeOption('default', 'Default'),
    ]),
    thinkingToggle: freezeOptions([
        makeOption('default', 'Default'),
        makeOption('off', 'Off'),
        makeOption('on', 'On'),
    ]),
    glm52: freezeOptions([
        makeOption('default', 'Default (Max)'),
        makeOption('off', 'Off'),
        makeOption('minimal', 'Minimal'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
        makeOption('xhigh', 'Extra high'),
        makeOption('max', 'Max'),
    ]),
    glm53: freezeOptions([
        makeOption('default', 'Default (Max)'),
        makeOption('low', 'Low'),
        makeOption('high', 'High'),
        makeOption('max', 'Max'),
    ]),
    forced: freezeOptions([
        makeOption('always', 'Always on'),
    ]),
    kimiK3: freezeOptions([
        makeOption('default', 'Default (Max)'),
        makeOption('low', 'Low'),
        makeOption('high', 'High'),
        makeOption('max', 'Max'),
    ]),
    minimaxM3: freezeOptions([
        makeOption('default', 'Default (Adaptive)'),
        makeOption('adaptive', 'Adaptive'),
        makeOption('off', 'Off'),
    ]),
    openaiGpt56: freezeOptions([
        makeOption('default', 'Default (Medium)'),
        makeOption('none', 'None'),
        makeOption('minimal', 'Minimal'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
        makeOption('xhigh', 'Extra high'),
        makeOption('max', 'Max'),
    ]),
    openaiGpt54: freezeOptions([
        makeOption('default', 'Default (Medium)'),
        makeOption('none', 'None'),
        makeOption('minimal', 'Minimal'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
        makeOption('xhigh', 'Extra high'),
    ]),
    openaiGpt51: freezeOptions([
        makeOption('default', 'Default (Medium)'),
        makeOption('none', 'None'),
        makeOption('minimal', 'Minimal'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
    ]),
    openaiGpt5: freezeOptions([
        makeOption('default', 'Default (Medium)'),
        makeOption('minimal', 'Minimal'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
    ]),
    openaiOSeries: freezeOptions([
        makeOption('default', 'Default (Medium)'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
    ]),
    claudeEffortXhigh: freezeOptions([
        makeOption('default', 'Default (High)'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
        makeOption('xhigh', 'Extra high'),
        makeOption('max', 'Max'),
    ]),
    claudeEffort: freezeOptions([
        makeOption('default', 'Default (High)'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low'),
        makeOption('medium', 'Medium'),
        makeOption('high', 'High'),
        makeOption('max', 'Max'),
    ]),
    claudeBudget: freezeOptions([
        makeOption('default', 'Default (Off)'),
        makeOption('off', 'Off'),
        makeOption('low', 'Low (4K tokens)'),
        makeOption('medium', 'Medium (8K tokens)'),
        makeOption('high', 'High (16K tokens)'),
    ]),
};

const TEMPERATURE_DEFAULT_HINT = 'Temperature is sent to this model when supported.';
const TEMPERATURE_UNSUPPORTED_HINT = 'Temperature is not supported by this model.';
const THINKING_TEMPERATURE_HINT = 'Temperature is unavailable while thinking is enabled.';

/**
 * Builds one immutable capability profile.
 *
 * @param {object} profile
 * @param {string} profile.id - Stable profile identifier
 * @param {string} profile.protocol - Wire-protocol key for buildThinkingRequest
 * @param {ReadonlyArray<{value: string, label: string, description?: string}>} profile.options
 * @param {string} profile.defaultLevel - Option value used when nothing applies
 * @param {boolean} [profile.temperatureSupported=true]
 * @param {string} [profile.temperatureHint]
 * @param {Object<string, boolean>} [profile.temperatureByLevel]
 * @param {Object<string, string>} [profile.aliases]
 * @param {string} [profile.hint]
 */
function makeProfile({
    id,
    protocol,
    options,
    defaultLevel,
    temperatureSupported = true,
    temperatureHint = TEMPERATURE_DEFAULT_HINT,
    temperatureByLevel,
    aliases = {},
    hint,
}) {
    return Object.freeze({
        id,
        protocol,
        options,
        defaultLevel,
        temperature: Object.freeze({
            supported: temperatureSupported,
            hint: temperatureHint,
        }),
        temperatureByLevel: temperatureByLevel
            ? Object.freeze({ ...temperatureByLevel })
            : undefined,
        aliases: Object.freeze({ ...aliases }),
        hint: hint || '',
    });
}

/** Generic behavior retained for unknown OpenAI-compatible endpoints. */
const GENERIC_PROFILE = makeProfile({
    id: 'generic',
    protocol: 'generic',
    options: GENERIC_OPTIONS,
    defaultLevel: 'default',
    hint: 'Uses the generic OpenAI-compatible reasoning_effort field.',
});

const OLLAMA_GPT_OSS_PROFILE = makeProfile({
    id: 'ollama-gpt-oss',
    protocol: 'ollama-reasoning-effort',
    options: PROFILE_OPTIONS.ollamaGptOss,
    defaultLevel: 'default',
    hint: 'GPT-OSS supports Low, Medium, and High reasoning; thinking cannot be fully disabled.',
});

const OLLAMA_QWEN_PROFILE = makeProfile({
    id: 'ollama-qwen-thinking',
    protocol: 'ollama-qwen-thinking',
    options: PROFILE_OPTIONS.ollamaQwen,
    defaultLevel: 'default',
    hint: 'Ollama Qwen models expose a thinking toggle and effort through the OpenAI-compatible endpoint.',
});

const QWEN3_VLLM_PROFILE = makeProfile({
    id: 'vllm-qwen3',
    protocol: 'vllm-qwen3-thinking',
    options: PROFILE_OPTIONS.qwen3Vllm,
    defaultLevel: 'default',
    hint: 'vLLM Qwen3 uses chat-template thinking and a per-request reasoning token budget.',
});

const DEEPSEEK_V4_PROFILE = makeProfile({
    id: 'deepseek-v4',
    protocol: 'deepseek-v4-thinking',
    options: PROFILE_OPTIONS.deepseekV4,
    defaultLevel: 'default',
    temperatureByLevel: { default: false, off: true, low: false, high: false, max: false },
    aliases: { on: 'high', medium: 'high', xhigh: 'high' },
    hint: 'DeepSeek V4 supports Off, Low, High, and Max. Temperature is unavailable while thinking.',
});

const DEEPSEEK_LEGACY_CHAT_PROFILE = makeProfile({
    id: 'deepseek-legacy-chat',
    protocol: 'none',
    options: PROFILE_OPTIONS.legacyDefault,
    defaultLevel: 'default',
    hint: 'Legacy DeepSeek Chat uses the model default and has no thinking dial.',
});

const DEEPSEEK_LEGACY_REASONER_PROFILE = makeProfile({
    id: 'deepseek-legacy-reasoner',
    protocol: 'none',
    options: PROFILE_OPTIONS.forced,
    defaultLevel: 'always',
    temperatureSupported: false,
    temperatureHint: TEMPERATURE_UNSUPPORTED_HINT,
    aliases: { default: 'always' },
    hint: 'Legacy DeepSeek Reasoner always thinks and does not accept Temperature.',
});

const GLM_TOGGLE_PROFILE = makeProfile({
    id: 'glm-thinking-toggle',
    protocol: 'glm-thinking-toggle',
    options: PROFILE_OPTIONS.thinkingToggle,
    defaultLevel: 'default',
    hint: 'GLM 4.5-4.7 exposes a thinking toggle rather than effort levels.',
});

const GLM_52_PROFILE = makeProfile({
    id: 'glm-5.2',
    protocol: 'glm-effort',
    options: PROFILE_OPTIONS.glm52,
    defaultLevel: 'default',
    aliases: { none: 'off' },
    hint: 'GLM-5.2 supports Minimal, Low, Medium, High, Extra high, Max, and Off.',
});

const GLM_53_PROFILE = makeProfile({
    id: 'glm-5.3',
    protocol: 'glm-forced-effort',
    options: PROFILE_OPTIONS.glm53,
    defaultLevel: 'default',
    aliases: { off: 'default' },
    hint: 'GLM-5.3 and GLM-5.3-Flash always think; only Low, High, and Max are available.',
});

const KIMI_K3_PROFILE = makeProfile({
    id: 'kimi-k3',
    protocol: 'kimi-reasoning-effort',
    options: PROFILE_OPTIONS.kimiK3,
    defaultLevel: 'default',
    temperatureSupported: false,
    temperatureHint: TEMPERATURE_UNSUPPORTED_HINT,
    hint: 'Kimi K3 always thinks and supports Low, High, and Max effort.',
});

const KIMI_K2_PROFILE = makeProfile({
    id: 'kimi-k2-thinking',
    protocol: 'kimi-thinking-toggle',
    options: PROFILE_OPTIONS.thinkingToggle,
    defaultLevel: 'default',
    temperatureSupported: false,
    temperatureHint: TEMPERATURE_UNSUPPORTED_HINT,
    hint: 'Kimi K2 exposes enabled or disabled thinking; Temperature is fixed by the model.',
});

const KIMI_CODE_PROFILE = makeProfile({
    id: 'kimi-k2-code',
    protocol: 'kimi-forced-thinking',
    options: PROFILE_OPTIONS.forced,
    defaultLevel: 'always',
    temperatureSupported: false,
    temperatureHint: TEMPERATURE_UNSUPPORTED_HINT,
    aliases: { default: 'always', off: 'always', on: 'always' },
    hint: 'Kimi K2.7 Code always thinks and does not expose Temperature or a thinking dial.',
});

const MINIMAX_M3_PROFILE = makeProfile({
    id: 'minimax-m3',
    protocol: 'minimax-m3-thinking',
    options: PROFILE_OPTIONS.minimaxM3,
    defaultLevel: 'default',
    hint: 'MiniMax M3 supports Adaptive thinking or a disabled mode; it has no effort scale.',
});

const MINIMAX_M2_PROFILE = makeProfile({
    id: 'minimax-m2',
    protocol: 'minimax-forced-thinking',
    options: PROFILE_OPTIONS.forced,
    defaultLevel: 'always',
    aliases: { default: 'always', off: 'always' },
    hint: 'MiniMax M2.x always thinks; thinking cannot be disabled.',
});

// OpenAI reasoning models reject Temperature while reasoning is active;
// it is accepted only when effort is `none` (openai-python #3073).
const OPENAI_TEMPERATURE_BY_LEVEL = Object.freeze({
    default: false, none: true, minimal: false, low: false,
    medium: false, high: false, xhigh: false, max: false,
});

const OPENAI_GPT_56_PROFILE = makeProfile({
    id: 'openai-gpt-5.6',
    protocol: 'openai-reasoning-effort',
    options: PROFILE_OPTIONS.openaiGpt56,
    defaultLevel: 'default',
    temperatureByLevel: OPENAI_TEMPERATURE_BY_LEVEL,
    aliases: { off: 'none' },
    hint: 'GPT-5.6 supports None through Max reasoning effort. Temperature is accepted only at None.',
});

const OPENAI_GPT_54_PROFILE = makeProfile({
    id: 'openai-gpt-5.4',
    protocol: 'openai-reasoning-effort',
    options: PROFILE_OPTIONS.openaiGpt54,
    defaultLevel: 'default',
    temperatureByLevel: OPENAI_TEMPERATURE_BY_LEVEL,
    aliases: { off: 'none' },
    hint: 'GPT-5.4/5.5 and Codex-Max support None through Extra high effort. Temperature is accepted only at None.',
});

const OPENAI_GPT_51_PROFILE = makeProfile({
    id: 'openai-gpt-5.1',
    protocol: 'openai-reasoning-effort',
    options: PROFILE_OPTIONS.openaiGpt51,
    defaultLevel: 'default',
    temperatureByLevel: OPENAI_TEMPERATURE_BY_LEVEL,
    aliases: { off: 'none' },
    hint: 'GPT-5.1 supports None, Minimal, Low, Medium, and High effort. Temperature is accepted only at None.',
});

const OPENAI_GPT_5_PROFILE = makeProfile({
    id: 'openai-gpt-5',
    protocol: 'openai-reasoning-effort',
    options: PROFILE_OPTIONS.openaiGpt5,
    defaultLevel: 'default',
    temperatureSupported: false,
    temperatureHint: 'GPT-5 reasoning models do not accept Temperature.',
    hint: 'GPT-5 supports Minimal, Low, Medium, and High reasoning effort; reasoning cannot be disabled on this generation.',
});

const OPENAI_O_SERIES_PROFILE = makeProfile({
    id: 'openai-o-series',
    protocol: 'openai-reasoning-effort',
    options: PROFILE_OPTIONS.openaiOSeries,
    defaultLevel: 'default',
    temperatureSupported: false,
    temperatureHint: TEMPERATURE_UNSUPPORTED_HINT,
    hint: 'o-series reasoning models support Low, Medium, and High effort; sampling parameters are fixed.',
});

const OPENAI_LEGACY_PROFILE = makeProfile({
    id: 'openai-legacy',
    protocol: 'none',
    options: PROFILE_OPTIONS.legacyDefault,
    defaultLevel: 'default',
    hint: 'This OpenAI model has no reasoning dial; only Temperature applies.',
});

// Claude effort era (4.6+): output_config.effort controls the whole token
// spend, thinking included; thinking can be disabled explicitly. Claude
// temperatures are 0-1 (the client clamps the generic 0-2 range).
const CLAUDE_EFFORT_XHIGH_PROFILE = makeProfile({
    id: 'claude-effort-xhigh',
    protocol: 'claude-effort',
    options: PROFILE_OPTIONS.claudeEffortXhigh,
    defaultLevel: 'default',
    hint: 'This Claude supports Off plus Low through Max effort (adaptive thinking). Temperature range is 0-1.',
});

const CLAUDE_EFFORT_PROFILE = makeProfile({
    id: 'claude-effort',
    protocol: 'claude-effort',
    options: PROFILE_OPTIONS.claudeEffort,
    defaultLevel: 'default',
    hint: 'This Claude supports Off plus Low, Medium, High, and Max effort (adaptive thinking). Temperature range is 0-1.',
});

const CLAUDE_BUDGET_PROFILE = makeProfile({
    id: 'claude-thinking-budget',
    protocol: 'claude-thinking-budget',
    options: PROFILE_OPTIONS.claudeBudget,
    defaultLevel: 'default',
    temperatureByLevel: { default: true, off: true, low: false, medium: false, high: false },
    hint: 'Claude 4.5 uses extended thinking with a token budget; Temperature is unavailable while thinking is enabled.',
});

const CLAUDE_LEGACY_PROFILE = makeProfile({
    id: 'claude-legacy',
    protocol: 'none',
    options: PROFILE_OPTIONS.legacyDefault,
    defaultLevel: 'default',
    hint: 'This Claude model has no thinking dial; only Temperature applies (range 0-1).',
});

const PROVIDER_ALIASES = Object.freeze({
    'minimax-cn': 'minimax',
});

function normalizeProvider(provider) {
    const value = String(provider || '').trim().toLowerCase();
    return PROVIDER_ALIASES[value] || value;
}

function normalizeModel(model) {
    return String(model || '').trim().toLowerCase();
}

// Model-specific profiles are explicit. Gateway providers can use these only
// when the selected model id names a known upstream model; unknown models stay
// on the generic profile and receive no vendor-only fields.
const PROFILE_MATCHERS = [
    {
        providers: ['ollama'],
        profile: OLLAMA_GPT_OSS_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_.]?oss(?:[^a-z]|$)/i,
    },
    {
        providers: ['ollama'],
        profile: OLLAMA_QWEN_PROFILE,
        matcher: /(?:^|[/:._-])qwen[-_.]?3(?:[^a-z]|$)/i,
    },
    {
        providers: ['vllm'],
        profile: QWEN3_VLLM_PROFILE,
        matcher: /(?:^|[/:._-])qwen[-_.]?3(?:[^a-z]|$)/i,
    },
    {
        providers: ['deepseek', 'custom', 'zhongkeyu'],
        profile: DEEPSEEK_V4_PROFILE,
        matcher: /(?:^|[/:._-])deepseek[-_. ]?v4[-_. ]?(?:flash|pro)(?:[^a-z]|$)/i,
    },
    {
        providers: ['deepseek'],
        profile: DEEPSEEK_LEGACY_REASONER_PROFILE,
        matcher: /(?:^|[/:._-])deepseek[-_. ]?reasoner(?:[^a-z]|$)/i,
    },
    {
        providers: ['deepseek'],
        profile: DEEPSEEK_LEGACY_CHAT_PROFILE,
        matcher: /(?:^|[/:._-])deepseek[-_. ]?chat(?:[^a-z]|$)/i,
    },
    {
        providers: ['glm', 'custom', 'zhongkeyu'],
        profile: GLM_53_PROFILE,
        matcher: /(?:^|[/:._-])glm[-_. ]?5[._-]?3(?:[-_. ]?flash)?(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['glm', 'custom', 'zhongkeyu'],
        profile: GLM_52_PROFILE,
        matcher: /(?:^|[/:._-])glm[-_. ]?5[._-]?2(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['glm', 'custom', 'zhongkeyu'],
        profile: GLM_TOGGLE_PROFILE,
        matcher: /(?:^|[/:._-])glm[-_. ]?4[._-]?(?:5|6|7)(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['kimi', 'custom', 'zhongkeyu'],
        profile: KIMI_K3_PROFILE,
        matcher: /(?:^|[/:._-])kimi[-_. ]?k3(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['kimi', 'custom', 'zhongkeyu'],
        profile: KIMI_K2_PROFILE,
        matcher: /(?:^|[/:._-])kimi[-_. ]?k2[._-]?6(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['kimi', 'custom', 'zhongkeyu'],
        profile: KIMI_CODE_PROFILE,
        matcher: /(?:^|[/:._-])kimi[-_. ]?k2[._-]?7(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['kimi', 'custom', 'zhongkeyu'],
        profile: KIMI_K2_PROFILE,
        matcher: /(?:^|[/:._-])kimi[-_. ]?k2(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['minimax', 'custom', 'zhongkeyu'],
        profile: MINIMAX_M3_PROFILE,
        matcher: /(?:^|[/:._-])minimax[-_. ]?m3(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['minimax', 'custom', 'zhongkeyu'],
        profile: MINIMAX_M2_PROFILE,
        matcher: /(?:^|[/:._-])minimax[-_. ]?m2(?:[^0-9a-z]|$)/i,
    },
    // OpenAI generations, newest first; gateways match these only when the
    // model id explicitly names a GPT/o-series upstream model.
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_GPT_54_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_. ]?5[._-]?1[-_. ]?codex[-_. ]?max(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_GPT_56_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_. ]?5[._-]?6(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_GPT_54_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_. ]?5[._-]?[45](?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_GPT_51_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_. ]?5[._-]?1(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_GPT_5_PROFILE,
        matcher: /(?:^|[/:._-])gpt[-_. ]?5(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['openai', 'custom', 'zhongkeyu'],
        profile: OPENAI_O_SERIES_PROFILE,
        matcher: /(?:^|[/:._-])o[134](?:[^0-9a-z]|$)/i,
    },
    // Claude generations. The wire fields are Anthropic-specific, so these
    // profiles never apply to gateway providers.
    {
        providers: ['claude'],
        profile: CLAUDE_BUDGET_PROFILE,
        matcher: /(?:^|[/:._-])(?:opus|sonnet|haiku)[-_. ]?4[._-]?5(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['claude'],
        profile: CLAUDE_EFFORT_XHIGH_PROFILE,
        matcher: /(?:^|[/:._-])(?:opus|sonnet)[-_. ]?4[._-]?[78](?:[^0-9a-z]|$)|(?:opus|sonnet|fable|mythos)[-_. ]?5(?:[^0-9a-z]|$)/i,
    },
    {
        providers: ['claude'],
        profile: CLAUDE_EFFORT_PROFILE,
        matcher: /(?:^|[/:._-])(?:opus|sonnet|haiku)[-_. ]?4[._-]?6(?:[^0-9a-z]|$)|mythos[-_. ]?preview/i,
    },
    {
        providers: ['claude'],
        profile: CLAUDE_LEGACY_PROFILE,
        matcher: /(?:^|[/:._-])claude[-_. ]?3(?:[^0-9a-z]|$)/i,
    },
];

/**
 * Provider-level fallbacks for first-party APIs whose unknown/newer models
 * share the provider's parameter family. Gateway and local providers keep
 * the generic profile for anything unrecognized.
 */
const PROVIDER_FALLBACK_PROFILES = Object.freeze({
    openai: OPENAI_LEGACY_PROFILE,
    claude: CLAUDE_EFFORT_PROFILE,
});

/**
 * Returns the capability profile for a provider/model pair.
 *
 * @param {string} provider - Provider/backend id
 * @param {string} model - Model id, including an optional namespace or tag
 * @returns {object} Immutable capability profile
 */
export function getModelCapabilities(provider, model) {
    const normalizedProvider = normalizeProvider(provider);
    const normalizedModel = normalizeModel(model);
    const match = PROFILE_MATCHERS.find((entry) => (
        entry.providers.includes(normalizedProvider) && entry.matcher.test(normalizedModel)
    ));
    if (match) return match.profile;
    return PROVIDER_FALLBACK_PROFILES[normalizedProvider] || GENERIC_PROFILE;
}

/**
 * Returns true when value is valid for the persisted canonical field.
 *
 * @param {unknown} value
 * @returns {boolean}
 */
export function isThinkingLevel(value) {
    return typeof value === 'string' && LEVEL_VALUES.has(value);
}

function hasOption(capabilities, value) {
    return capabilities.options.some((candidate) => candidate.value === value);
}

/**
 * Resolves a persisted/UI value to one of a profile's visible options.
 * Unsupported old values use a profile alias when one is safe; otherwise the
 * profile default is used.
 *
 * @param {object} capabilities - Result from getModelCapabilities
 * @param {string} value - Persisted or UI thinking value
 * @returns {string} A value present in capabilities.options
 */
export function resolveThinkingLevel(capabilities, value) {
    const profile = capabilities && Array.isArray(capabilities.options) ? capabilities : GENERIC_PROFILE;
    const candidate = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (hasOption(profile, candidate)) return candidate;

    const alias = profile.aliases[candidate];
    if (alias && hasOption(profile, alias)) return alias;

    if (hasOption(profile, profile.defaultLevel)) return profile.defaultLevel;
    return profile.options[0].value;
}

/**
 * Maps a canonical level to documented request fields for a capability profile.
 * Temperature is handled separately because support can depend on the level.
 *
 * @param {object} capabilities - Result from getModelCapabilities
 * @param {string} value - Persisted or UI thinking value
 * @returns {object} Protocol fields to merge into a chat-completions body
 */
export function buildThinkingRequest(capabilities, value) {
    const profile = capabilities && Array.isArray(capabilities.options) ? capabilities : GENERIC_PROFILE;
    const requested = typeof value === 'string' ? value.trim().toLowerCase() : '';
    const level = resolveThinkingLevel(profile, value);

    // `default` leaves the provider/model's own default behavior in control.
    if (!requested || requested === 'default') return {};

    switch (profile.protocol) {
    case 'generic':
    case 'ollama-reasoning-effort':
        if (level === 'off') return { reasoning_effort: 'none' };
        return ['low', 'medium', 'high', 'max'].includes(level)
            ? { reasoning_effort: level }
            : {};
    case 'ollama-qwen-thinking':
        if (level === 'off') return { reasoning_effort: 'none' };
        return ['low', 'medium', 'high'].includes(level)
            ? { reasoning_effort: level }
            : {};
    case 'vllm-qwen3-thinking':
        if (level === 'off') return { chat_template_kwargs: { enable_thinking: false } };
        return {
            chat_template_kwargs: { enable_thinking: true },
            thinking_token_budget: { low: 4096, medium: 8192, high: 16384 }[level] || 8192,
        };
    case 'deepseek-v4-thinking':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return ['low', 'high', 'max'].includes(level)
            ? { thinking: { type: 'enabled' }, reasoning_effort: level }
            : {};
    case 'glm-thinking-toggle':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return level === 'on' ? { thinking: { type: 'enabled' } } : {};
    case 'glm-effort':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return ['minimal', 'low', 'medium', 'high', 'xhigh', 'max'].includes(level)
            ? { thinking: { type: 'enabled' }, reasoning_effort: level }
            : {};
    case 'glm-forced-effort':
        return ['low', 'high', 'max'].includes(level)
            ? { thinking: { type: 'enabled' }, reasoning_effort: level }
            : {};
    case 'kimi-reasoning-effort':
        return ['low', 'high', 'max'].includes(level)
            ? { reasoning_effort: level }
            : {};
    case 'kimi-thinking-toggle':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return level === 'on' ? { thinking: { type: 'enabled' } } : {};
    case 'kimi-forced-thinking':
        return {};
    case 'minimax-m3-thinking':
        return level === 'off'
            ? { thinking: { type: 'disabled' }, reasoning_split: true }
            : level === 'adaptive'
                ? { thinking: { type: 'adaptive' }, reasoning_split: true }
                : {};
    case 'minimax-forced-thinking':
        return level === 'always' ? { reasoning_split: true } : {};
    case 'openai-reasoning-effort':
        // Canonical values are wire-identical for OpenAI (none/minimal/…/max).
        return ['none', 'minimal', 'low', 'medium', 'high', 'xhigh', 'max'].includes(level)
            ? { reasoning_effort: level }
            : {};
    case 'claude-effort':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return ['low', 'medium', 'high', 'xhigh', 'max'].includes(level)
            ? { output_config: { effort: level } }
            : {};
    case 'claude-thinking-budget':
        if (level === 'off') return { thinking: { type: 'disabled' } };
        return ['low', 'medium', 'high'].includes(level)
            ? { thinking: { type: 'enabled', budget_tokens: { low: 4096, medium: 8192, high: 16384 }[level] } }
            : {};
    case 'none':
    default:
        return {};
    }
}

/**
 * Returns whether Temperature can be sent for a profile/level.
 *
 * @param {object} capabilities - Result from getModelCapabilities
 * @param {string} value - Persisted or UI thinking value
 * @returns {boolean}
 */
export function isTemperatureSupported(capabilities, value) {
    const profile = capabilities && Array.isArray(capabilities.options) ? capabilities : GENERIC_PROFILE;
    const level = resolveThinkingLevel(profile, value);
    if (profile.temperature.supported === false) return false;
    if (profile.temperatureByLevel && Object.prototype.hasOwnProperty.call(profile.temperatureByLevel, level)) {
        return profile.temperatureByLevel[level] !== false;
    }
    return true;
}

/**
 * Returns a user-facing Temperature explanation for a profile/level.
 *
 * @param {object} capabilities - Result from getModelCapabilities
 * @param {string} value - Persisted or UI thinking value
 * @returns {string}
 */
export function getTemperatureHint(capabilities, value) {
    const profile = capabilities && Array.isArray(capabilities.options) ? capabilities : GENERIC_PROFILE;
    if (!isTemperatureSupported(profile, value)) {
        if (profile.temperatureByLevel && profile.temperatureByLevel[resolveThinkingLevel(profile, value)] === false) {
            return THINKING_TEMPERATURE_HINT;
        }
        return profile.temperature.hint || TEMPERATURE_UNSUPPORTED_HINT;
    }
    return profile.temperature.hint || TEMPERATURE_DEFAULT_HINT;
}

/**
 * Returns the profile's explanatory thinking hint.
 *
 * @param {object} capabilities - Result from getModelCapabilities
 * @returns {string}
 */
export function getThinkingHint(capabilities) {
    const profile = capabilities && Array.isArray(capabilities.options) ? capabilities : GENERIC_PROFILE;
    return profile.hint;
}
