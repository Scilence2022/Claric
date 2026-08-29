/**
 * Tests for scripts/generate-manifest.cjs (CommonJS module).
 * Covers XML escaping, template rendering, GUID resolution/persistence,
 * and version normalization from package.json.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const {
  generateManifest,
  renderTemplate,
  escapeXml,
  resolveGuid,
  buildAppDomainsBlock,
} = require('../scripts/generate-manifest.cjs');

const TEMPLATE = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  <Id>\${GUID}</Id>
  <Version>\${VERSION}</Version>
  <DisplayName DefaultValue="\${DISPLAY_NAME}"/>
  <SupportUrl DefaultValue="\${SUPPORT_URL}"/>
  \${APP_DOMAINS_BLOCK}</OfficeApp>
`;

/**
 * Creates an isolated project dir with template + package.json.
 * Returns the dir path (caller must clean up).
 */
function makeProjectDir({ version = '1.2.3' } = {}) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'manifest-test-'));
  fs.writeFileSync(path.join(dir, 'manifest.template.xml'), TEMPLATE);
  fs.writeFileSync(path.join(dir, 'package.json'), JSON.stringify({ version }));
  return dir;
}

describe('generate-manifest.cjs', () => {
  let projectDir;

  beforeEach(() => {
    projectDir = makeProjectDir();
  });

  afterEach(() => {
    fs.rmSync(projectDir, { recursive: true, force: true });
  });

  describe('escapeXml', () => {
    test('escapes the five XML special characters', () => {
      expect(escapeXml(`<a>&"b'</a>`)).toBe('&lt;a&gt;&amp;&quot;b&apos;&lt;/a&gt;');
    });
  });

  describe('renderTemplate', () => {
    test('substitutes known placeholders and escapes their values', () => {
      const out = renderTemplate('<u>${PROTOCOL}://${HOST}</u>', {
        PROTOCOL: 'https',
        HOST: 'host&name',
      });
      expect(out).toBe('<u>https://host&amp;name</u>');
    });

    test('leaves unknown placeholders untouched', () => {
      expect(renderTemplate('<x>${UNKNOWN}</x>', {})).toBe('<x>${UNKNOWN}</x>');
    });
  });

  describe('resolveGuid', () => {
    test('generates and persists a stable GUID', () => {
      const first = resolveGuid(projectDir, null);
      expect(first).toMatch(/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/);
      expect(resolveGuid(projectDir, null)).toBe(first);
      expect(fs.readFileSync(path.join(projectDir, '.manifest-guid'), 'utf8').trim()).toBe(first);
    });

    test('env GUID wins over the persisted file', () => {
      resolveGuid(projectDir, null);
      const pinned = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
      expect(resolveGuid(projectDir, pinned)).toBe(pinned);
    });

    test('invalid env GUID falls back to persisted/generated GUID', () => {
      expect(resolveGuid(projectDir, 'not-a-guid')).toMatch(/^[0-9a-f]{8}-/);
    });
  });

  describe('generateManifest', () => {
    test('HOST_PORT takes precedence over PORT for manifest URLs', () => {
        const oldEnv = { ...process.env };
        process.env.HOST_PORT = '4123';
        process.env.PORT = '3000';
        try {
            const outPath = generateManifest({ rootDir: projectDir });
            const xml = fs.readFileSync(outPath, 'utf8');
            expect(xml).toContain('localhost:4123/');
        } finally {
            process.env = oldEnv;
        }
    });

    test('writes a manifest with a real GUID, 4-part version, and URLs', () => {
      const outPath = generateManifest({ rootDir: projectDir });
      expect(outPath).toBe(path.join(projectDir, 'manifest.xml'));

      const xml = fs.readFileSync(outPath, 'utf8');
      expect(xml).toMatch(/<Id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}<\/Id>/);
      expect(xml).toContain('<Version>1.2.3.0</Version>');
      expect(xml).toContain('https://localhost:3000/');
      expect(xml).not.toContain('${');
    });

    test('regeneration keeps the same add-in identity', () => {
      generateManifest({ rootDir: projectDir });
      const first = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8').match(/<Id>([^<]+)</)[1];
      generateManifest({ rootDir: projectDir });
      const second = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8').match(/<Id>([^<]+)</)[1];
      expect(second).toBe(first);
    });

    test('throws a clear error when the template is missing', () => {
      fs.rmSync(path.join(projectDir, 'manifest.template.xml'));
      expect(() => generateManifest({ rootDir: projectDir })).toThrow(/Missing manifest template/);
    });
  });
});

describe('store identity generation', () => {
  let projectDir;

  beforeEach(() => {
    projectDir = makeProjectDir();
  });

  afterEach(() => {
    fs.rmSync(projectDir, { recursive: true, force: true });
  });

  function withEnv(env, fn) {
    const old = { ...process.env };
    Object.assign(process.env, env);
    try {
      return fn();
    } finally {
      process.env = old;
    }
  }

  test('buildAppDomainsBlock renders an escaped AppDomains element', () => {
    expect(buildAppDomainsBlock(null)).toBe('');
    expect(buildAppDomainsBlock('')).toBe('');
    expect(buildAppDomainsBlock('api.example.com, llm.example.org')).toBe(
      '<AppDomains>\n' +
      '    <AppDomain>api.example.com</AppDomain>\n' +
      '    <AppDomain>llm.example.org</AppDomain>\n' +
      '  </AppDomains>\n  '
    );
    // Env values must not be able to break out of the element.
    expect(buildAppDomainsBlock('a</AppDomain><script>')).toContain('&lt;/AppDomain&gt;');
  });

  test('APP_DOMAINS env produces the AppDomains element; unset omits it entirely', () => {
    withEnv({ APP_DOMAINS: 'api.example.com' }, () => {
      generateManifest({ rootDir: projectDir });
      const xml = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8');
      expect(xml).toContain('<AppDomains>');
      expect(xml).toContain('<AppDomain>api.example.com</AppDomain>');
      expect(xml).not.toContain('${APP_DOMAINS_BLOCK}');
    });
    withEnv({}, () => {
      generateManifest({ rootDir: projectDir });
      const xml = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8');
      expect(xml).not.toContain('<AppDomains>');
      expect(xml).not.toContain('${');
    });
  });

  test('SUPPORT_URL env overrides the support page', () => {
    withEnv({ SUPPORT_URL: 'https://github.com/Scilence2022/Claric/issues' }, () => {
      generateManifest({ rootDir: projectDir });
      const xml = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8');
      expect(xml).toContain('<SupportUrl DefaultValue="https://github.com/Scilence2022/Claric/issues"/>');
    });
  });

  test('SUPPORT_URL falls back to the add-in root when unset', () => {
    withEnv({}, () => {
      generateManifest({ rootDir: projectDir });
      const xml = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8');
      expect(xml).toContain('<SupportUrl DefaultValue="https://localhost:3000/"/>');
    });
  });

  test('DISPLAY_NAME env overrides the display name with a store-style suffix', () => {
    withEnv({ DISPLAY_NAME: 'Claric — AI Redlining for Word' }, () => {
      generateManifest({ rootDir: projectDir });
      const xml = fs.readFileSync(path.join(projectDir, 'manifest.xml'), 'utf8');
      expect(xml).toContain('<DisplayName DefaultValue="Claric — AI Redlining for Word"/>');
    });
  });
});
