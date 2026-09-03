const makeWebpackConfig = require('../webpack.config.cjs');

const {
  parseAllowedHosts,
  parseCorsOrigins,
  resolveCorsOrigin,
  corsHeaders,
  parseProxyTarget,
  buildLlmProxies,
} = makeWebpackConfig.__testing;

describe('webpack dev-server host and CORS policy', () => {
  test('defaults to loopback hostnames and never accepts the all sentinel', () => {
    expect(parseAllowedHosts()).toEqual(['localhost', '127.0.0.1', '[::1]']);
    expect(parseAllowedHosts(' dev.example, localhost, all, ALL '))
      .toEqual(['dev.example', 'localhost']);
  });

  test('defaults CORS to local HTTPS origins and parses explicit origins', () => {
    expect(parseCorsOrigins(undefined, 3000)).toEqual([
      'https://localhost:3000',
      'https://127.0.0.1:3000',
      'https://[::1]:3000',
    ]);
    expect(parseCorsOrigins('https://dev.example:444, javascript:alert(1)', 3000))
      .toEqual(['https://dev.example:444']);
    expect(parseCorsOrigins('*', 3000)).toEqual(['*']);
  });

  test('does not reflect an untrusted request origin', () => {
    const allowed = ['https://localhost:3000'];
    expect(resolveCorsOrigin('https://localhost:3000', allowed)).toBe('https://localhost:3000');
    expect(resolveCorsOrigin('https://evil.example', allowed)).toBeNull();

    expect(corsHeaders({ headers: { origin: 'https://localhost:3000' } }, allowed))
      .toMatchObject({
        'Access-Control-Allow-Origin': 'https://localhost:3000',
        Vary: 'Origin',
      });
    expect(corsHeaders({ headers: { origin: 'https://evil.example' } }, allowed))
      .not.toHaveProperty('Access-Control-Allow-Origin');
  });

  test('wildcard CORS remains an explicit opt-in only', () => {
    expect(corsHeaders({ headers: { origin: 'https://example.com' } }, ['*']))
      .toHaveProperty('Access-Control-Allow-Origin', '*');
  });

  test('generated config uses an array allowlist and dynamic CORS headers', () => {
    const config = makeWebpackConfig({}, { mode: 'development' });
    expect(Array.isArray(config.devServer.allowedHosts)).toBe(true);
    expect(config.devServer.allowedHosts).not.toContain('all');
    expect(typeof config.devServer.headers).toBe('function');
  });
});

describe('webpack dev-server proxy target policy', () => {
  test.each([
    'https://api.example.com',
    'http://localhost:11434',
    'http://127.0.0.1:8026',
    'http://[::1]:8026',
    'http://host.docker.internal:11434',
  ])('allows %s', (target) => {
    expect(parseProxyTarget(target)).toBeInstanceOf(URL);
  });

  test.each([
    'http://remote.example/v1',
    'http://localhost.evil/v1',
    'ftp://localhost/model',
    'javascript:alert(1)',
    'https://user:password@example.com/v1',
    '//remote.example/v1',
  ])('rejects %s', (target) => {
    expect(parseProxyTarget(target)).toBeNull();
  });

  test('proxy hooks restrict CORS and retain local HTTP backends', () => {
    const proxies = buildLlmProxies({
      DEV_SERVER_PORT: 3000,
      DEV_SERVER_CORS_ORIGINS: ['https://localhost:3000'],
      OLLAMA_PROXY_PATH: '/ollama',
      OLLAMA_PROXY_TARGET: 'http://localhost:11434',
    });
    const proxy = proxies[0];
    const headers = {};
    const response = {
      setHeader: (name, value) => { headers[name] = value; },
      end: jest.fn(),
      statusCode: 0,
    };

    try {
      expect(proxies).toHaveLength(1);
      expect(proxy.target).toBe('http://localhost:11434/');
      expect(proxy.bypass({
        method: 'OPTIONS',
        headers: { origin: 'https://localhost:3000' },
      }, response)).toBe(true);
      expect(headers['Access-Control-Allow-Origin']).toBe('https://localhost:3000');

      const upstreamHeaders = { 'access-control-allow-origin': '*' };
      proxy.onProxyRes(
        { headers: upstreamHeaders, statusCode: 200 },
        { url: '/ollama/v1/models', headers: { origin: 'https://evil.example' } },
      );
      expect(upstreamHeaders['access-control-allow-origin']).toBeUndefined();
      expect(upstreamHeaders['Access-Control-Allow-Origin']).toBeUndefined();
    } finally {
      proxy.agent.destroy();
    }
  });
});
