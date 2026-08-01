import { describe, expect, it } from 'vitest';
import {
  normalizeMediaReplaceRequest,
  resolveMediaReplacementSource,
} from './media-replace.internal.js';

describe('media source replacement inputs', () => {
  it('normalizes and resolves embedded and external sources with creation semantics', async () => {
    const request = normalizeMediaReplaceRequest(
      'audio',
      'data:audio/mpeg;base64,AQID',
      { contentType: 'audio/mpeg', fileName: 'voice.mp3' },
    );
    const resolved = await resolveMediaReplacementSource(request);
    expect(resolved).toEqual({
      type: 'embedded',
      bytes: Uint8Array.of(1, 2, 3),
      contentType: 'audio/mpeg',
      extension: '.mp3',
    });
    expect(await resolveMediaReplacementSource(normalizeMediaReplaceRequest(
      'video',
      'https://example.com/video.mp4',
    ))).toEqual({ type: 'external', url: 'https://example.com/video.mp4' });
  });

  it('accepts detached transcoding and rejects replacement-only invalid state', async () => {
    const bytes = Uint8Array.of(1);
    const request = normalizeMediaReplaceRequest('audio', bytes, {
      transcode: async () => ({
        bytes: Uint8Array.of(2, 3),
        contentType: 'audio/wav',
        extension: '.wav',
      }),
    });
    bytes[0] = 9;
    expect(await resolveMediaReplacementSource(request)).toMatchObject({
      bytes: Uint8Array.of(2, 3),
      contentType: 'audio/wav',
      extension: '.wav',
    });

    let reads = 0;
    const accessor = Object.defineProperty({}, 'contentType', {
      enumerable: true,
      get() { reads += 1; return 'audio/mpeg'; },
    });
    expect(() => normalizeMediaReplaceRequest('audio', Uint8Array.of(1), accessor)).toThrow(/data property/);
    expect(reads).toBe(0);
    for (const options of [null, [], new Date(), { poster: Uint8Array.of(1) }, { x: 1 }]) {
      expect(() => normalizeMediaReplaceRequest('audio', Uint8Array.of(1), options as never)).toThrow();
    }
    await expect(resolveMediaReplacementSource(normalizeMediaReplaceRequest(
      'audio',
      'https://example.com/audio.mp3',
      { transcode: async (value, contentType) => ({ bytes: value, contentType }) },
    ))).rejects.toThrow(/External media cannot be transcoded/);
  });
});
