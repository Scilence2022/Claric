/**
 * Icon Generator
 *
 * Renders the Claric mark (assets/icon.svg) to the PNG sizes the Office
 * manifest references (16/32/64/80) plus 128 for HiDPI, writing them to
 * assets/icon-{size}.png. Runs as part of `npm run build` BEFORE webpack's
 * CopyWebpackPlugin copies assets/ into dist/, so a production build always
 * ships current icons (and the Docker builder stage does the same).
 *
 * The generator is best-effort: if `sharp` (a build-time devDependency) is
 * unavailable (e.g. `npm ci --omit=dev`), it warns and leaves any existing
 * PNGs in place rather than failing the build.
 */

const fs = require('fs');
const path = require('path');

const rootDir = path.resolve(__dirname, '..');
const svgPath = path.join(rootDir, 'assets', 'icon.svg');
const ICON_SIZES = [16, 32, 64, 80, 128];

/** Renders the SVG source to a PNG buffer at a square size. */
async function renderPng(sharp, size) {
  const svg = fs.readFileSync(svgPath);
  return sharp(svg, { density: 288 }).resize(size, size).png().toBuffer();
}

async function main() {
  if (!fs.existsSync(svgPath)) {
    console.log('generate-icons: assets/icon.svg not found — skipping.');
    return;
  }

  let sharp;
  try {
    sharp = require('sharp');
  } catch {
    console.log('generate-icons: sharp is not installed — keeping existing icon PNGs.');
    return;
  }

  for (const size of ICON_SIZES) {
    const outPath = path.join(rootDir, 'assets', `icon-${size}.png`);
    const png = await renderPng(sharp, size);
    fs.writeFileSync(outPath, png);
    console.log(`generate-icons: wrote assets/icon-${size}.png (${png.length} bytes)`);
  }
}

main().catch((err) => {
  console.error(`generate-icons failed: ${err.message}`);
  console.error('Keeping the previously committed icon PNGs.');
  process.exitCode = 1;
});

module.exports = { ICON_SIZES };
