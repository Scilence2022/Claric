/**
 * Unit tests for src/lib/platform.js
 */

const { getHostPlatform, supportsTrackedRowOps } = require('../src/lib/platform.js');

describe('getHostPlatform', () => {
  test('returns Office.context.platform when present', () => {
    expect(getHostPlatform({ context: { platform: 'PC' } })).toBe('PC');
    expect(getHostPlatform({ context: { platform: 'OfficeOnline' } })).toBe('OfficeOnline');
  });

  test('returns "unknown" when Office is absent or the read fails', () => {
    expect(getHostPlatform(undefined)).toBe('unknown');
    expect(getHostPlatform({})).toBe('unknown');
    expect(getHostPlatform({
      get context() { throw new Error('no context'); },
    })).toBe('unknown');
  });
});

describe('supportsTrackedRowOps', () => {
  test('desktop hosts track table row revisions', () => {
    expect(supportsTrackedRowOps('PC')).toBe(true);
    expect(supportsTrackedRowOps('Mac')).toBe(true);
    expect(supportsTrackedRowOps('Universal')).toBe(true);
  });

  test('web/mobile/unknown hosts do not', () => {
    expect(supportsTrackedRowOps('OfficeOnline')).toBe(false);
    expect(supportsTrackedRowOps('iOS')).toBe(false);
    expect(supportsTrackedRowOps('Android')).toBe(false);
    expect(supportsTrackedRowOps('unknown')).toBe(false);
  });
});
