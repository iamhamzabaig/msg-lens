import { describe, it, expect } from 'vitest';
import { parseMsgFile } from '../../src/index';

describe('parseMsgFile', () => {
  it('returns INVALID_CFB for empty buffer', () => {
    const result = parseMsgFile(new ArrayBuffer(0));
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('INVALID_CFB');
    }
  });

  it('returns INVALID_CFB for random data', () => {
    const junk = new Uint8Array(1024);
    for (let i = 0; i < junk.length; i++) junk[i] = i % 256;
    const result = parseMsgFile(junk);
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error.code).toBe('INVALID_CFB');
    }
  });

  it('accepts Uint8Array input', () => {
    const result = parseMsgFile(new Uint8Array(0));
    expect(result.success).toBe(false);
  });

  it('never throws on any input', () => {
    // Various malformed inputs — none should throw
    expect(() => parseMsgFile(new ArrayBuffer(0))).not.toThrow();
    expect(() => parseMsgFile(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]))).not.toThrow();
    expect(() => parseMsgFile(new Uint8Array(10000))).not.toThrow();
  });
});
