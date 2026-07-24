export interface CodecOwnership {
  readonly elements?: readonly string[];
  readonly relationshipTypes?: readonly string[];
  readonly contentTypes?: readonly string[];
}

export interface DecodeContext {
  readonly partUri: string;
  readonly value: unknown;
}

export interface EncodeContext extends DecodeContext {}
export interface ValidationContext extends DecodeContext {}

export interface CodecDiagnostic {
  readonly severity: 'info' | 'warning' | 'error';
  readonly code: string;
  readonly message: string;
  readonly partUri?: string;
}

export interface XmlPatch {
  readonly partUri: string;
  readonly start: number;
  readonly end: number;
  readonly replacement: string;
}

export interface FeatureCodec<T> {
  readonly id: string;
  readonly priority?: number;
  readonly ownership: CodecOwnership;
  detect(context: DecodeContext): boolean;
  decode(context: DecodeContext): T;
  encode(value: T, context: EncodeContext): readonly XmlPatch[];
  validate(value: T, context: ValidationContext): readonly CodecDiagnostic[];
}

export interface CodecRegistration {
  readonly id: string;
  readonly priority?: number;
  readonly ownership: CodecOwnership;
}

export class CodecOwnershipError extends Error {
  constructor(readonly firstCodec: string, readonly secondCodec: string, readonly resource: string) {
    super(`Codec ownership conflict for ${resource}: ${firstCodec} and ${secondCodec}`);
    this.name = 'CodecOwnershipError';
  }
}

export class CodecRegistry {
  readonly #codecs = new Map<string, CodecRegistration>();

  register<T>(codec: CodecRegistration | FeatureCodec<T>): void {
    if (this.#codecs.has(codec.id)) throw new Error(`Codec ${codec.id} is already registered`);
    for (const existing of this.#codecs.values()) this.assertNoConflict(existing, codec);
    this.#codecs.set(codec.id, codec);
  }

  get codecs(): readonly CodecRegistration[] {
    return [...this.#codecs.values()].sort((left, right) => (right.priority ?? 100) - (left.priority ?? 100));
  }

  private assertNoConflict(left: CodecRegistration, right: CodecRegistration): void {
    if ((left.priority ?? 100) !== (right.priority ?? 100)) return;
    for (const key of ['elements', 'relationshipTypes', 'contentTypes'] as const) {
      const rightValues = new Set(right.ownership[key] ?? []);
      const conflict = (left.ownership[key] ?? []).find((value) => rightValues.has(value));
      if (conflict) throw new CodecOwnershipError(left.id, right.id, `${key}:${conflict}`);
    }
  }
}
