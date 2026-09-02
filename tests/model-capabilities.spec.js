/**
 * Model-specific thinking profiles and protocol mapping contracts.
 */

const {
  THINKING_LEVEL_VALUES,
  getModelCapabilities,
  resolveThinkingLevel,
  buildThinkingRequest,
  isTemperatureSupported,
  getThinkingHint,
  getTemperatureHint,
} = require('../src/lib/model-capabilities.js');

describe('model capabilities', () => {
  test('exports the canonical thinking values', () => {
    expect(THINKING_LEVEL_VALUES).toEqual(expect.arrayContaining([
      'default', 'off', 'on', 'adaptive', 'always', 'low', 'medium', 'high',
      'xhigh', 'max', 'minimal', 'none',
    ]));
  });

  test.each([
    ['ollama', 'gpt-oss:20b', 'ollama-gpt-oss', ['default', 'low', 'medium', 'high']],
    ['vllm', 'Qwen/Qwen3.5-35B-A3B', 'vllm-qwen3', ['default', 'off', 'low', 'medium', 'high']],
    ['deepseek', 'deepseek-v4-flash', 'deepseek-v4', ['default', 'off', 'low', 'high', 'max']],
    ['glm', 'glm-4.7', 'glm-thinking-toggle', ['default', 'off', 'on']],
    ['glm', 'glm-5.2', 'glm-5.2', ['default', 'off', 'minimal', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['glm', 'glm-5.3-flash', 'glm-5.3', ['default', 'low', 'high', 'max']],
    ['kimi', 'moonshotai/Kimi-K3', 'kimi-k3', ['default', 'low', 'high', 'max']],
    ['kimi', 'kimi-k2.6', 'kimi-k2-thinking', ['default', 'off', 'on']],
    ['kimi', 'kimi-k2.7-code:latest', 'kimi-k2-code', ['always']],
    ['minimax-cn', 'MiniMax-M3', 'minimax-m3', ['default', 'adaptive', 'off']],
    ['minimax', 'MiniMax-M2.5', 'minimax-m2', ['always']],
  ])('matches %s/%s', (provider, model, id, values) => {
    const profile = getModelCapabilities(provider, model);
    expect(profile.id).toBe(id);
    expect(profile.options.map((option) => option.value)).toEqual(values);
  });

  test('unknown models retain generic behavior, including known provider ids', () => {
    const profile = getModelCapabilities('zhongkeyu', 'some-new-model');
    expect(profile.id).toBe('generic');
    expect(profile.options.map((option) => option.value)).toEqual(['default', 'low', 'medium', 'high']);
    expect(buildThinkingRequest(profile, 'default')).toEqual({});
    expect(buildThinkingRequest(profile, 'high')).toEqual({ reasoning_effort: 'high' });
  });

  test('explicit upstream model ids can use a gateway profile', () => {
    expect(getModelCapabilities('zhongkeyu', 'glm-5.3-flash').id).toBe('glm-5.3');
    expect(getModelCapabilities('custom', 'deepseek-v4-pro').id).toBe('deepseek-v4');
  });

  test('resolves unsupported values to a safe profile value', () => {
    const qwen = getModelCapabilities('vllm', 'qwen3.5-35b');
    expect(resolveThinkingLevel(qwen, 'default')).toBe('default');
    // 'on' is not a vLLM Qwen level and has no safe alias: profile default wins.
    expect(resolveThinkingLevel(qwen, 'on')).toBe('default');

    const forced = getModelCapabilities('kimi', 'kimi-k2.7-code');
    expect(resolveThinkingLevel(forced, 'off')).toBe('always');

    const deepseek = getModelCapabilities('deepseek', 'deepseek-v4-flash');
    expect(resolveThinkingLevel(deepseek, 'medium')).toBe('high');
    expect(resolveThinkingLevel(deepseek, 'xhigh')).toBe('high');

    // GLM-5.3 always thinks: Off resolves away instead of being sent.
    const glm53 = getModelCapabilities('glm', 'glm-5.3');
    expect(resolveThinkingLevel(glm53, 'off')).toBe('default');
  });

  test('default never sends a thinking field (provider default stays in control)', () => {
    for (const [provider, model] of [
      ['ollama', 'gpt-oss:20b'],
      ['ollama', 'qwen3:32b'],
      ['vllm', 'qwen3.5-35b-a3b'],
      ['deepseek', 'deepseek-v4-pro'],
      ['deepseek', 'deepseek-reasoner'],
      ['glm', 'glm-4.5'],
      ['glm', 'glm-5.2'],
      ['glm', 'glm-5.3-flash'],
      ['kimi', 'kimi-k3'],
      ['kimi', 'kimi-k2.6'],
      ['minimax', 'MiniMax-M3'],
      ['custom', 'my-model'],
    ]) {
      expect(buildThinkingRequest(getModelCapabilities(provider, model), 'default')).toEqual({});
    }
  });

  test('maps Ollama GPT-OSS levels to reasoning_effort', () => {
    const profile = getModelCapabilities('ollama', 'gpt-oss:20b');
    expect(buildThinkingRequest(profile, 'low')).toEqual({ reasoning_effort: 'low' });
    expect(buildThinkingRequest(profile, 'high')).toEqual({ reasoning_effort: 'high' });
    // GPT-OSS cannot fully disable thinking; there is no Off option.
    expect(profile.options.some((o) => o.value === 'off')).toBe(false);
  });

  test('maps Ollama Qwen levels, including the thinking toggle off', () => {
    const profile = getModelCapabilities('ollama', 'qwen3:32b');
    expect(buildThinkingRequest(profile, 'off')).toEqual({ reasoning_effort: 'none' });
    expect(buildThinkingRequest(profile, 'medium')).toEqual({ reasoning_effort: 'medium' });
  });

  test('maps vLLM Qwen levels to chat template and token budget', () => {
    const profile = getModelCapabilities('vllm', 'qwen3.5-35b-a3b');
    expect(buildThinkingRequest(profile, 'off')).toEqual({
      chat_template_kwargs: { enable_thinking: false },
    });
    expect(buildThinkingRequest(profile, 'low')).toEqual({
      chat_template_kwargs: { enable_thinking: true },
      thinking_token_budget: 4096,
    });
    expect(buildThinkingRequest(profile, 'medium')).toEqual({
      chat_template_kwargs: { enable_thinking: true },
      thinking_token_budget: 8192,
    });
    expect(buildThinkingRequest(profile, 'high')).toEqual({
      chat_template_kwargs: { enable_thinking: true },
      thinking_token_budget: 16384,
    });
  });

  test('maps DeepSeek V4 thinking and effort without unsupported medium/xhigh', () => {
    const profile = getModelCapabilities('deepseek', 'deepseek-v4-pro');
    expect(buildThinkingRequest(profile, 'off')).toEqual({ thinking: { type: 'disabled' } });
    expect(buildThinkingRequest(profile, 'low')).toEqual({
      thinking: { type: 'enabled' },
      reasoning_effort: 'low',
    });
    expect(buildThinkingRequest(profile, 'max')).toEqual({
      thinking: { type: 'enabled' },
      reasoning_effort: 'max',
    });
    expect(isTemperatureSupported(profile, 'off')).toBe(true);
    expect(isTemperatureSupported(profile, 'low')).toBe(false);
    expect(isTemperatureSupported(profile, 'high')).toBe(false);
    expect(isTemperatureSupported(profile, 'default')).toBe(false);
  });

  test('legacy DeepSeek models have no thinking dial', () => {
    const chat = getModelCapabilities('deepseek', 'deepseek-chat');
    expect(chat.id).toBe('deepseek-legacy-chat');
    expect(buildThinkingRequest(chat, 'default')).toEqual({});

    const reasoner = getModelCapabilities('deepseek', 'deepseek-reasoner');
    expect(reasoner.id).toBe('deepseek-legacy-reasoner');
    expect(resolveThinkingLevel(reasoner, 'off')).toBe('always');
    expect(isTemperatureSupported(reasoner, 'always')).toBe(false);
  });

  test('maps GLM protocol variants across generations', () => {
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-4.5'), 'on'))
      .toEqual({ thinking: { type: 'enabled' } });
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-4.5'), 'off'))
      .toEqual({ thinking: { type: 'disabled' } });
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-5.2'), 'minimal'))
      .toEqual({ thinking: { type: 'enabled' }, reasoning_effort: 'minimal' });
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-5.2'), 'xhigh'))
      .toEqual({ thinking: { type: 'enabled' }, reasoning_effort: 'xhigh' });
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-5.2'), 'off'))
      .toEqual({ thinking: { type: 'disabled' } });
    expect(buildThinkingRequest(getModelCapabilities('glm', 'glm-5.3'), 'low'))
      .toEqual({ thinking: { type: 'enabled' }, reasoning_effort: 'low' });
  });

  test('maps Kimi K3 effort, K2 toggle, and the always-thinking Code profile', () => {
    expect(buildThinkingRequest(getModelCapabilities('kimi', 'kimi-k3'), 'max'))
      .toEqual({ reasoning_effort: 'max' });
    expect(buildThinkingRequest(getModelCapabilities('kimi', 'kimi-k2.6'), 'off'))
      .toEqual({ thinking: { type: 'disabled' } });
    expect(buildThinkingRequest(getModelCapabilities('kimi', 'kimi-k2.6'), 'on'))
      .toEqual({ thinking: { type: 'enabled' } });
    expect(buildThinkingRequest(getModelCapabilities('kimi', 'kimi-k2.7-code'), 'always')).toEqual({});
  });

  test('maps MiniMax M3 and keeps M2 thinking enabled', () => {
    expect(buildThinkingRequest(getModelCapabilities('minimax', 'MiniMax-M3'), 'off'))
      .toEqual({ thinking: { type: 'disabled' }, reasoning_split: true });
    expect(buildThinkingRequest(getModelCapabilities('minimax', 'MiniMax-M3'), 'adaptive'))
      .toEqual({ thinking: { type: 'adaptive' }, reasoning_split: true });
    expect(buildThinkingRequest(getModelCapabilities('minimax', 'MiniMax-M2.1'), 'always'))
      .toEqual({ reasoning_split: true });
  });

  test('reports temperature restrictions and hints', () => {
    const k26 = getModelCapabilities('kimi', 'kimi-k2.6');
    expect(isTemperatureSupported(k26, 'off')).toBe(false);
    expect(getTemperatureHint(k26, 'off')).toContain('not supported');

    const generic = getModelCapabilities('custom', 'my-openai-model');
    expect(isTemperatureSupported(generic, 'default')).toBe(true);
    expect(getTemperatureHint(generic, 'default')).toContain('sent');

    const deepseek = getModelCapabilities('deepseek', 'deepseek-v4-flash');
    expect(getTemperatureHint(deepseek, 'high')).toContain('unavailable while thinking');
    expect(getTemperatureHint(deepseek, 'off')).toContain('sent');
  });

  test('exposes a user-facing thinking hint for every known profile', () => {
    for (const [provider, model] of [
      ['ollama', 'gpt-oss:20b'],
      ['vllm', 'qwen3.5-35b'],
      ['deepseek', 'deepseek-v4-flash'],
      ['glm', 'glm-5.2'],
      ['kimi', 'kimi-k3'],
      ['minimax', 'MiniMax-M3'],
      ['openai', 'gpt-5.1'],
      ['claude', 'claude-sonnet-4-6'],
    ]) {
      expect(getThinkingHint(getModelCapabilities(provider, model)).length).toBeGreaterThan(0);
    }
  });
});

describe('OpenAI model capabilities', () => {
  test.each([
    ['gpt-5.6', 'openai-gpt-5.6', ['default', 'none', 'minimal', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['gpt-5.4', 'openai-gpt-5.4', ['default', 'none', 'minimal', 'low', 'medium', 'high', 'xhigh']],
    ['gpt-5.1-codex-max', 'openai-gpt-5.4', ['default', 'none', 'minimal', 'low', 'medium', 'high', 'xhigh']],
    ['gpt-5.1', 'openai-gpt-5.1', ['default', 'none', 'minimal', 'low', 'medium', 'high']],
    ['gpt-5', 'openai-gpt-5', ['default', 'minimal', 'low', 'medium', 'high']],
    ['gpt-5-mini-2025-08-07', 'openai-gpt-5', ['default', 'minimal', 'low', 'medium', 'high']],
    ['o3', 'openai-o-series', ['default', 'low', 'medium', 'high']],
    ['o4-mini', 'openai-o-series', ['default', 'low', 'medium', 'high']],
    ['gpt-4o', 'openai-legacy', ['default']],
    ['gpt-4.1-mini', 'openai-legacy', ['default']],
  ])('matches %s', (model, id, values) => {
    const profile = getModelCapabilities('openai', model);
    expect(profile.id).toBe(id);
    expect(profile.options.map((option) => option.value)).toEqual(values);
  });

  test('reasoning effort is wire-identical and none disables reasoning', () => {
    const profile = getModelCapabilities('openai', 'gpt-5.6');
    expect(buildThinkingRequest(profile, 'none')).toEqual({ reasoning_effort: 'none' });
    expect(buildThinkingRequest(profile, 'minimal')).toEqual({ reasoning_effort: 'minimal' });
    expect(buildThinkingRequest(profile, 'xhigh')).toEqual({ reasoning_effort: 'xhigh' });
    expect(buildThinkingRequest(profile, 'max')).toEqual({ reasoning_effort: 'max' });
    expect(buildThinkingRequest(profile, 'default')).toEqual({});
  });

  test('the canonical off value aliases to none', () => {
    const profile = getModelCapabilities('openai', 'gpt-5.1');
    expect(resolveThinkingLevel(profile, 'off')).toBe('none');
    expect(buildThinkingRequest(profile, 'off')).toEqual({ reasoning_effort: 'none' });
  });

  test('temperature is accepted only at none for reasoning models', () => {
    const profile = getModelCapabilities('openai', 'gpt-5.1');
    expect(isTemperatureSupported(profile, 'none')).toBe(true);
    expect(isTemperatureSupported(profile, 'low')).toBe(false);
    expect(isTemperatureSupported(profile, 'default')).toBe(false);
    // GPT-5.0 has no none level, so it never accepts temperature.
    const gpt5 = getModelCapabilities('openai', 'gpt-5');
    expect(isTemperatureSupported(gpt5, 'minimal')).toBe(false);
    // o-series fix sampling entirely.
    const o3 = getModelCapabilities('openai', 'o3');
    expect(isTemperatureSupported(o3, 'low')).toBe(false);
  });

  test('legacy models keep temperature and have no reasoning dial', () => {
    const legacy = getModelCapabilities('openai', 'gpt-4o');
    expect(isTemperatureSupported(legacy, 'default')).toBe(true);
    expect(buildThinkingRequest(legacy, 'high')).toEqual({});
  });

  test('unknown models on the openai provider fall back to the legacy profile', () => {
    expect(getModelCapabilities('openai', 'some-future-model').id).toBe('openai-legacy');
  });

  test('gateways pick up GPT profiles only for explicit GPT model ids', () => {
    expect(getModelCapabilities('zhongkeyu', 'gpt-5.1').id).toBe('openai-gpt-5.1');
    expect(getModelCapabilities('custom', 'openai/gpt-5.6').id).toBe('openai-gpt-5.6');
    expect(getModelCapabilities('zhongkeyu', 'some-new-model').id).toBe('generic');
  });
});

describe('Claude model capabilities', () => {
  test.each([
    ['claude-opus-4-7', 'claude-effort-xhigh', ['default', 'off', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['claude-opus-4-8-20260715', 'claude-effort-xhigh', ['default', 'off', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['claude-sonnet-5', 'claude-effort-xhigh', ['default', 'off', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['claude-fable-5', 'claude-effort-xhigh', ['default', 'off', 'low', 'medium', 'high', 'xhigh', 'max']],
    ['claude-sonnet-4-6', 'claude-effort', ['default', 'off', 'low', 'medium', 'high', 'max']],
    ['claude-opus-4-6', 'claude-effort', ['default', 'off', 'low', 'medium', 'high', 'max']],
    ['claude-opus-4-5', 'claude-thinking-budget', ['default', 'off', 'low', 'medium', 'high']],
    ['claude-haiku-4-5-20251001', 'claude-thinking-budget', ['default', 'off', 'low', 'medium', 'high']],
    ['claude-3-5-sonnet-20241022', 'claude-legacy', ['default']],
  ])('matches %s', (model, id, values) => {
    const profile = getModelCapabilities('claude', model);
    expect(profile.id).toBe(id);
    expect(profile.options.map((option) => option.value)).toEqual(values);
  });

  test('effort levels map to output_config.effort', () => {
    const profile = getModelCapabilities('claude', 'claude-opus-4-7');
    expect(buildThinkingRequest(profile, 'low')).toEqual({ output_config: { effort: 'low' } });
    expect(buildThinkingRequest(profile, 'xhigh')).toEqual({ output_config: { effort: 'xhigh' } });
    expect(buildThinkingRequest(profile, 'max')).toEqual({ output_config: { effort: 'max' } });
    expect(buildThinkingRequest(profile, 'default')).toEqual({});
    expect(buildThinkingRequest(profile, 'off')).toEqual({ thinking: { type: 'disabled' } });
  });

  test('xhigh is only offered on the models that support it', () => {
    const effort = getModelCapabilities('claude', 'claude-sonnet-4-6');
    expect(resolveThinkingLevel(effort, 'xhigh')).toBe('default');
    expect(buildThinkingRequest(effort, 'xhigh')).toEqual({});
  });

  test('budget-era models map levels to thinking budgets', () => {
    const profile = getModelCapabilities('claude', 'claude-opus-4-5');
    expect(buildThinkingRequest(profile, 'low')).toEqual({ thinking: { type: 'enabled', budget_tokens: 4096 } });
    expect(buildThinkingRequest(profile, 'high')).toEqual({ thinking: { type: 'enabled', budget_tokens: 16384 } });
    expect(buildThinkingRequest(profile, 'off')).toEqual({ thinking: { type: 'disabled' } });
    expect(isTemperatureSupported(profile, 'low')).toBe(false);
    expect(isTemperatureSupported(profile, 'off')).toBe(true);
  });

  test('adaptive-era models accept temperature at every effort level', () => {
    const profile = getModelCapabilities('claude', 'claude-sonnet-5');
    expect(isTemperatureSupported(profile, 'xhigh')).toBe(true);
    expect(isTemperatureSupported(profile, 'off')).toBe(true);
  });

  test('unknown models on the claude provider fall back to the effort profile', () => {
    expect(getModelCapabilities('claude', 'claude-future-9').id).toBe('claude-effort');
  });

  test('claude profiles never leak onto gateway providers', () => {
    expect(getModelCapabilities('zhongkeyu', 'claude-opus-4-7').id).toBe('generic');
  });
});
