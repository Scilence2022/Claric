const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const dotenv = require('dotenv');

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Loads manifest-related environment variables (with .env support).
 *
 * HOST_PORT takes precedence over PORT: manifest URLs must reference the
 * host-visible port, while a container may listen on a different internal
 * port (docker-compose publishes HOST_PORT -> 3000).
 *
 * @param {string} rootDir - Project root (used to locate .env)
 * @returns {{ HOST: string, PORT: string, PROTOCOL: string, ADDIN_GUID: string|null }}
 */
function getEnv(rootDir) {
  dotenv.config({ path: path.join(rootDir, '.env') });
  return {
    HOST: process.env.HOST || 'localhost',
    PORT: process.env.HOST_PORT || process.env.PORT || '3000',
    PROTOCOL: process.env.PROTOCOL || 'https',
    ADDIN_GUID: process.env.ADDIN_GUID || null,
  };
}

/**
 * Escapes XML special characters so env values cannot break the manifest.
 *
 * @param {string} value
 * @returns {string}
 */
function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Substitutes ${NAME} placeholders in the manifest template.
 * Unknown placeholders are left untouched so they are easy to spot
 * in validation errors rather than silently producing empty values.
 *
 * @param {string} template
 * @param {Record<string, string>} values
 * @returns {string}
 */
function renderTemplate(template, values) {
  return template.replace(/\$\{(\w+)\}/g, (match, name) =>
    Object.prototype.hasOwnProperty.call(values, name)
      ? escapeXml(values[name])
      : match
  );
}

/**
 * Resolves the add-in GUID.
 *
 * Word identifies an add-in by its manifest GUID. A stable GUID is required:
 * a new GUID on every restart makes Word treat the add-in as newly installed
 * (trust prompts and per-add-in state reset each time). Resolution order:
 *
 *   1. ADDIN_GUID environment variable (explicit deployment pinning)
 *   2. Previously generated GUID persisted in <rootDir>/.manifest-guid
 *   3. A freshly generated random UUID (persisted for future runs)
 *
 * @param {string} rootDir
 * @param {string|null} envGuid
 * @returns {string}
 */
function resolveGuid(rootDir, envGuid) {
  if (envGuid && UUID_RE.test(envGuid)) {
    return envGuid.toLowerCase();
  }

  const guidPath = path.join(rootDir, '.manifest-guid');
  if (fs.existsSync(guidPath)) {
    const persisted = fs.readFileSync(guidPath, 'utf8').trim();
    if (UUID_RE.test(persisted)) {
      return persisted.toLowerCase();
    }
  }

  const generated = crypto.randomUUID();
  fs.writeFileSync(guidPath, generated, 'utf8');
  return generated;
}

/**
 * Reads the add-in version from package.json and converts it to the
 * four-part form required by the Office manifest schema (1.2.3 -> 1.2.3.0).
 *
 * @param {string} rootDir
 * @returns {string}
 */
function getVersion(rootDir) {
  const pkg = JSON.parse(fs.readFileSync(path.join(rootDir, 'package.json'), 'utf8'));
  const parts = String(pkg.version || '0.0.0').split('.');
  while (parts.length < 4) parts.push('0');
  return parts.slice(0, 4).join('.');
}

/**
 * Computes a short content hash of the add-in icon PNG, used as a cache-busting
 * query on the manifest's IconUrl / HighResolutionIconUrl. Word loads the
 * icon through WebView2's HTTP cache keyed by URL — the icon URL was
 * previously identical across deployments, so after changing the icon Word
 * kept serving the stale (placeholder/default) image. Baking a content hash
 * into the URL makes the URL change whenever the icon changes, forcing a
 * fresh fetch. Falls back to the package version when the file is missing.
 *
 * @param {string} rootDir
 * @returns {string}
 */
function getIconCache(rootDir) {
  for (const name of ['icon-64.png', 'icon-32.png']) {
    const iconPath = path.join(rootDir, 'assets', name);
    if (fs.existsSync(iconPath)) {
      return crypto.createHash('sha256').update(fs.readFileSync(iconPath)).digest('hex').slice(0, 10);
    }
  }
  return getVersion(rootDir);
}

/**
 * Generates manifest.xml from manifest.template.xml.
 *
 * Placeholders: ${HOST}, ${PORT}, ${PROTOCOL}, ${VERSION}, ${GUID},
 * ${DISPLAY_NAME}, ${ICON_CACHE}. All substituted values are XML-escaped.
 *
 * @param {object} [options]
 * @param {string} [options.rootDir] - Project root override (defaults to repo root)
 * @param {string} [options.displayName] - Display name override
 * @returns {string} Path of the generated manifest
 */
function generateManifest(options = {}) {
  const rootDir = options.rootDir || path.resolve(__dirname, '..');
  const templatePath = path.join(rootDir, 'manifest.template.xml');
  const outputPath = path.join(rootDir, 'manifest.xml');

  if (!fs.existsSync(templatePath)) {
    throw new Error(`Missing manifest template: ${templatePath}`);
  }

  const env = getEnv(rootDir);
  const template = fs.readFileSync(templatePath, 'utf8');

  // Preserve a real UUID already baked into a customized template so teams
  // that pin their GUID there keep it; otherwise resolve/persist one.
  const templateGuidMatch = template.match(/<Id>([^<]+)<\/Id>/);
  const templateGuid = templateGuidMatch && UUID_RE.test(templateGuidMatch[1].trim())
    ? templateGuidMatch[1].trim()
    : null;

  const values = {
    HOST: env.HOST,
    PORT: env.PORT,
    PROTOCOL: env.PROTOCOL,
    VERSION: getVersion(rootDir),
    GUID: templateGuid || resolveGuid(rootDir, env.ADDIN_GUID),
    DISPLAY_NAME: options.displayName || 'Claric',
    ICON_CACHE: getIconCache(rootDir),
  };

  const output = renderTemplate(template, values);
  fs.writeFileSync(outputPath, output, 'utf8');
  return outputPath;
}

if (require.main === module) {
  const out = generateManifest();
  console.log(`Manifest written to ${out}`);
}

module.exports = {
  generateManifest,
  renderTemplate,
  escapeXml,
  resolveGuid,
};
