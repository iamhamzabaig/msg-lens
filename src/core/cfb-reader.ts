import CFB from 'cfb';

export interface CfbStream {
  path: string;
  name: string;
  content: Uint8Array;
}

/**
 * Parse a raw .msg buffer into a CFB container and extract all streams.
 * @param data - Raw .msg file bytes
 * @returns Parsed streams grouped by path, or null if invalid
 */
export function parseCfb(data: ArrayBuffer | Uint8Array): CfbStream[] | null {
  try {
    const buf = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
    const cfb = CFB.parse(buf, { type: 'array' });

    const streams: CfbStream[] = [];
    for (let i = 0; i < cfb.FullPaths.length; i++) {
      const entry = cfb.FileIndex[i];
      // Only include streams (type 2), skip storages
      if (entry.type !== 2) continue;

      const content = entry.content instanceof Uint8Array
        ? entry.content
        : new Uint8Array(entry.content as number[]);

      streams.push({
        path: cfb.FullPaths[i],
        name: entry.name,
        content,
      });
    }

    return streams;
  } catch {
    return null;
  }
}

/**
 * Find all streams matching a path prefix.
 */
export function findStreams(streams: CfbStream[], prefix: string): CfbStream[] {
  return streams.filter(s => s.path.startsWith(prefix));
}

/**
 * Find a single stream by exact path.
 */
export function findStream(streams: CfbStream[], path: string): CfbStream | undefined {
  return streams.find(s => s.path === path);
}
