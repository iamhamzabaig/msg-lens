import { describe, it, expect } from 'vitest';
import { parseCfb, findStream, findStreams } from '../../src/core';

describe('parseCfb', () => {
  it('returns null for invalid/empty data', () => {
    expect(parseCfb(new ArrayBuffer(0))).toBeNull();
  });

  it('returns null for random bytes', () => {
    const junk = new Uint8Array(512);
    for (let i = 0; i < junk.length; i++) junk[i] = Math.floor(Math.random() * 256);
    expect(parseCfb(junk)).toBeNull();
  });

  it('returns null for too-small data', () => {
    expect(parseCfb(new Uint8Array([0xd0, 0xcf]))).toBeNull();
  });
});

describe('findStream / findStreams', () => {
  const streams = [
    { path: 'Root Entry/__substg1.0_0037001F', name: '__substg1.0_0037001F', content: new Uint8Array([]) },
    { path: 'Root Entry/__recip_version1.0_#00000000/__substg1.0_3001001F', name: '__substg1.0_3001001F', content: new Uint8Array([]) },
    { path: 'Root Entry/__recip_version1.0_#00000001/__substg1.0_3001001F', name: '__substg1.0_3001001F', content: new Uint8Array([]) },
  ];

  it('finds stream by exact path', () => {
    const result = findStream(streams, 'Root Entry/__substg1.0_0037001F');
    expect(result).toBeDefined();
    expect(result!.name).toBe('__substg1.0_0037001F');
  });

  it('returns undefined for missing stream', () => {
    expect(findStream(streams, 'Root Entry/missing')).toBeUndefined();
  });

  it('finds streams by prefix', () => {
    const recips = findStreams(streams, 'Root Entry/__recip');
    expect(recips).toHaveLength(2);
  });
});
