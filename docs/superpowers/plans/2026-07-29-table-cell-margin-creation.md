# Table Cell Margin Creation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. The user explicitly requires inline execution without subagents.

**Goal:** Let native `slide.addTable()` create cells with point-based direct four-side margins through `{ text, options: { margin } }`, while preserving canonical default bytes and the existing whole-replacement editor.

**Architecture:** Harden the shared `TextBoxMarginInput` normalizer to descriptor-safe detached reads, add one table-cell creation attribute renderer that overlays explicit sides on canonical narrow defaults, and extend normalized table cells with optional margin overrides. Keep existing read/edit ownership, transaction flow, and PptxGenJS adapter semantics unchanged.

**Tech Stack:** TypeScript, Vitest, lossless OOXML kernel, OPC transactions, PptxGenJS 4.0.1, tsup, npm tarballs, repository JSON CLI, LibreOffice, Poppler.

## Global Constraints

- Public creation property is `AddTableCellOptions.margin?: TextBoxMarginInput` and all values use points.
- Scalar applies to all sides; tuple order is T/R/B/L; named input overlays only supplied creation sides.
- Creation starts from top/bottom 3.6pt and left/right 7.2pt; omitted/undefined/empty named margin keeps current bytes.
- `setCellMargins()` remains a whole-replacement editor; partial named input clears missing direct sides there.
- Quantized raw margin values must fit signed Int32 and may be zero, negative, positive, or fractional points.
- Shared tuple/named normalization is descriptor-safe, getter-free, ordinary/null-prototype-only, and detached.
- OOXML creation attribute order remains marL/marR/marT/marB; borders remain L/R/T/B children followed by fill.
- PptxGenJS dual-unit runtime is documented and tested, not copied into the native API.
- No table-level margin, alignment/direction/fit creation, merge, hyperlink, rich text, auto-page, layout, or unrelated option work enters this slice.
- Every successful small item is reviewed, committed, pushed, and verified `origin/main...HEAD = 0 0`; QA-only gates do not create empty commits.
- Never add, stage, modify, or remove `.pnpm-store/`; stage only named target files.
- Use no subagents.

---

### Task 1: Make shared margin normalization descriptor-safe

**Files:**
- Modify: `packages/model/src/text-box-margins.internal.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: existing `TextBoxMarginInput`, `TextBoxMargins`, `normalizeTextBoxMargins()`, text-box margin creation/editing, and table-cell margin editing.
- Produces: the same normalized point values for valid inputs, with getter-free strict tuple/named object reads.

- [ ] **Step 1: Add failing descriptor-safety tests to the existing table-cell margin editor test**

Add these runtime inputs before the invalid-input loop:

```ts
let marginAccessorCalls = 0;
const accessorTuple = [1, 2, 3, 4];
Object.defineProperty(accessorTuple, '0', {
  get() {
    marginAccessorCalls += 1;
    return 1;
  },
  enumerable: true,
  configurable: true,
});
const accessorNamed = { right: 2 };
Object.defineProperty(accessorNamed, 'top', {
  get() {
    marginAccessorCalls += 1;
    return 1;
  },
  enumerable: true,
  configurable: true,
});
class MarginClass {
  top = 1;
}
const inheritedMargin = Object.create({ top: 1 });
const symbolMargin = { top: 1, [Symbol('margin')]: 2 };
const arraySubclass = new (class extends Array<number> {})(1, 2, 3, 4);
```

Pass all six values to `table.setCellMargins()` and require throws, `marginAccessorCalls === 0`, exact target slide bytes, and unchanged mutation journal. Add valid null-prototype named input:

```ts
const nullMargins = Object.assign(Object.create(null), {
  top: 1.500004,
  left: -2,
});
table.setCellMargins(0, 0, nullMargins);
expect(table.rows[0]!.cells[0]!.margins).toEqual({ top: 1.5, left: -2 });
```

Also retain existing scalar, ordinary tuple/object, clear, no-op, malformed repair, and rollback assertions so legal behavior cannot regress.

- [ ] **Step 2: Run the focused model test and confirm red**

```sh
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts \
  -t "table-cell margin"
```

Expected: an accessor getter is invoked or a class/inherited/exotic input is accepted.

- [ ] **Step 3: Rewrite tuple normalization using own descriptors**

Add a helper in `text-box-margins.internal.ts`:

```ts
function normalizeTuple(value: unknown[], context: string): TextBoxMargins {
  if (Object.getPrototypeOf(value) !== Array.prototype) {
    throw new TypeError(`${context} tuple must be an ordinary array`);
  }
  const length = Object.getOwnPropertyDescriptor(value, 'length');
  if (!length || !Object.hasOwn(length, 'value') || length.value !== 4) {
    throw new RangeError(`${context} tuple must contain exactly four values`);
  }
  const allowed = new Set(['0', '1', '2', '3', 'length']);
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !allowed.has(key)) {
      throw new TypeError(`${context} tuple contains unsupported property ${String(key)}`);
    }
  }
  const values = Array.from({ length: 4 }, (_, index) => {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(index));
    if (!descriptor) throw new TypeError(`${context} tuple must not contain sparse values`);
    if (!Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} tuple must contain only data values`);
    }
    return descriptor.value;
  });
  return {
    top: normalizeSide(values[0], `${context} top`),
    right: normalizeSide(values[1], `${context} right`),
    bottom: normalizeSide(values[2], `${context} bottom`),
    left: normalizeSide(values[3], `${context} left`),
  };
}
```

Call this helper immediately after `Array.isArray(value)`; do not read `value.length` or any index directly.

- [ ] **Step 4: Rewrite named normalization through a strict data-object copy**

Add:

```ts
function readMarginObject(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError(`${context} must be a number, four-value tuple, or margin object`);
  }
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) {
    throw new TypeError(`${context} must be an ordinary object`);
  }
  const result: Record<string, unknown> = {};
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || !SIDE_NAMES.has(key as keyof TextBoxMargins)) {
      throw new TypeError(`${context} contains unsupported property ${String(key)}`);
    }
    const descriptor = Object.getOwnPropertyDescriptor(value, key);
    if (!descriptor || !Object.hasOwn(descriptor, 'value')) {
      throw new TypeError(`${context} property ${key} must be a data property`);
    }
    result[key] = descriptor.value;
  }
  return result;
}
```

Iterate `SIDES`, check `Object.hasOwn(candidate, side)`, skip descriptor values that are `undefined`, and call `normalizeSide()` on every supplied number. Keep scalar normalization and signed-Int32 quantization unchanged.

- [ ] **Step 5: Run typecheck and all shared-normalizer consumers**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
```

Expected: existing text-box/table-cell margins and new descriptor-safety assertions pass.

- [ ] **Step 6: Review, commit, push, and prove synchronization**

```sh
git diff --check
git diff -- packages/model/src/text-box-margins.internal.ts packages/model/src/model.test.ts
git add -- packages/model/src/text-box-margins.internal.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: harden margin input normalization"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`; `.pnpm-store/` remains untracked and unstaged.

---

### Task 2: Normalize and render cell margins during native table creation

**Files:**
- Modify: `packages/model/src/table-cell-margins.internal.ts`
- Modify: `packages/model/src/table-create.internal.ts`
- Modify: `packages/model/src/table-create.internal.test.ts`

**Interfaces:**
- Consumes: Task 1 `normalizeTextBoxMargins()`.
- Produces:

```ts
export function renderTableCellMarginAttributes(
  margins: TextBoxMargins | undefined,
): string;
```

The renderer overlays explicit normalized sides on canonical top/bottom 3.6pt and left/right 7.2pt, then writes marL/marR/marT/marB.

- [ ] **Step 1: Add internal normalization and detachment tests first**

Add a table-creation test with these cells:

```ts
const sourceMargins = { top: 1.500004, left: -2 };
const nullMargins = Object.assign(Object.create(null), {
  right: 5,
  bottom: 6,
});
const rows = [[
  'String',
  { text: 'Empty options', options: {} },
  { text: 'Undefined margin', options: { margin: undefined } },
  { text: 'Empty margin', options: { margin: {} } },
  { text: 'Zero', options: { margin: 0 } },
  { text: 'Tuple', options: { margin: [1, 2, 3, 4] } },
  { text: 'Named', options: { margin: sourceMargins } },
  { text: 'Null prototype', options: { margin: nullMargins } },
  { text: 'Combined', options: {
    margin: { top: 4, left: 8 },
    border: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
    fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
  } },
]];
```

Require normalized cells to store `{}` only for explicit empty margin, four normalized sides for scalar/tuple, and explicit normalized overrides only for named values. Mutate `sourceMargins` and `nullMargins` after normalization and prove the definition remains `{ top: 1.5, left: -2 }` and `{ right: 5, bottom: 6 }`.

- [ ] **Step 2: Add exact default-overlay and XML-order assertions**

Build equivalent one-cell definitions for string, `{ text }`, empty options, `margin: undefined`, `margin: {}`, and `{ top: undefined, right: undefined }`; require byte-identical `renderTableGraphicFrame()` strings.

Require exact attributes:

```ts
expect(defaultXml).toContain(
  '<a:tcPr marL="91440" marR="91440" marT="45720" marB="45720">',
);
expect(zeroXml).toContain(
  '<a:tcPr marL="0" marR="0" marT="0" marB="0">',
);
expect(tupleXml).toContain(
  '<a:tcPr marL="50800" marR="25400" marT="12700" marB="38100">',
);
expect(namedXml).toContain(
  '<a:tcPr marL="-25400" marR="91440" marT="19050" marB="45720">',
);
expect(combinedXml).toMatch(
  /<a:tcPr marL="101600" marR="91440" marT="50800" marB="45720"><a:lnL[\s\S]*<\/a:lnB><a:solidFill>/,
);
```

- [ ] **Step 3: Extend the invalid creation matrix**

Add accessor options/margin cases without invoking getters:

```ts
let marginGetterCalls = 0;
const accessorMargin = { left: 1 };
Object.defineProperty(accessorMargin, 'top', {
  get() {
    marginGetterCalls += 1;
    return 2;
  },
  enumerable: true,
});
const accessorTuple = [1, 2, 3, 4];
Object.defineProperty(accessorTuple, '2', {
  get() {
    marginGetterCalls += 1;
    return 3;
  },
  enumerable: true,
});
```

Reject accessor/class/inherited/symbol named objects; accessor/sparse/wrong-length/extra-key/array-subclass tuples; null/boolean/string/object side values; `NaN`, infinities, and values whose quantized EMU exceeds signed Int32. Require `marginGetterCalls === 0`.

- [ ] **Step 4: Run the internal suite and confirm red**

```sh
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts
```

Expected: `margin` is rejected as an unsupported cell option.

- [ ] **Step 5: Add the canonical attribute renderer**

In `table-cell-margins.internal.ts`, add:

```ts
const DEFAULT_MARGINS: Required<TextBoxMargins> = {
  top: 3.6,
  right: 7.2,
  bottom: 3.6,
  left: 7.2,
};

export function renderTableCellMarginAttributes(
  margins: TextBoxMargins | undefined,
): string {
  return SIDES.map(([side, attribute]) => {
    const points = margins?.[side] ?? DEFAULT_MARGINS[side];
    return ` ${attribute}="${Math.round(points * EMU_PER_POINT)}"`;
  }).join('');
}
```

`SIDES` already uses left/right/top/bottom and marL/marR/marT/marB, so output remains canonical. Do not change read/replace editor functions.

- [ ] **Step 6: Extend normalized cells and render the attributes**

In `table-create.internal.ts`, import `normalizeTextBoxMargins`, `renderTableCellMarginAttributes`, and `TextBoxMargins`. Extend:

```ts
interface NormalizedTableCell {
  readonly text: string;
  readonly borders?: TableCellBorders;
  readonly fill?: TableCellFill;
  readonly margins?: TextBoxMargins;
}
```

Allow `['border', 'fill', 'margin']`, normalize margin after border/fill, and return it only when not `undefined`:

```ts
const margins = normalizeTextBoxMargins(
  options.margin as TextBoxMarginInput | undefined,
  `${context} margin`,
);
return {
  ...(borders === undefined ? {} : { borders }),
  ...(fill === undefined ? {} : { fill }),
  ...(margins === undefined ? {} : { margins }),
};
```

Import `type TextBoxMarginInput` with `TextBoxMargins`. Remove `CELL_MARGIN_HORIZONTAL` / `CELL_MARGIN_VERTICAL`, then render:

```ts
const marginAttributes = renderTableCellMarginAttributes(cell.margins);
return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr${marginAttributes}>${borders}${fill}</a:tcPr></a:tc>`;
```

- [ ] **Step 7: Run focused internal and editor regression suites**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/table-create.internal.test.ts \
  packages/model/src/model.test.ts
```

Expected: default bytes, overlay semantics, invalid getter matrix, and existing border/fill/margin editors pass.

- [ ] **Step 8: Review, commit, push, and prove synchronization**

```sh
git diff --check
git add -- \
  packages/model/src/table-cell-margins.internal.ts \
  packages/model/src/table-create.internal.ts \
  packages/model/src/table-create.internal.test.ts
git diff --cached --check
git commit -m "feat: render table cell margins during creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0` and only `?? .pnpm-store/` outside the committed item.

---

### Task 3: Expose the public option and prove model lifecycle

**Files:**
- Modify: `packages/model/src/slide.ts`
- Modify: `packages/model/src/model.test.ts`

**Interfaces:**
- Consumes: Task 2 internal margin creation.
- Produces public `AddTableCellOptions.margin?: TextBoxMarginInput`.

- [ ] **Step 1: Add public typing and lifecycle assertions first**

Import `type TextBoxMarginInput` through the public model surface and type mutable inputs:

```ts
const sourceNamed = { top: 4, left: 8 };
const sourceTuple: TextBoxMarginInput = [1, 2, 3, 4];
const sourceOptions: AddTableCellOptions = {
  margin: sourceNamed,
  border: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
};
```

Create a matrix containing omitted, zero, scalar, tuple, partial named, empty named, and margin+border+fill cells. Require immediate snapshots such as:

```ts
expect(table.rows[0]!.cells[0]!.margins).toEqual({
  top: 4,
  right: 7.2,
  bottom: 3.6,
  left: 8,
});
expect(table.rows[0]!.cells[1]!.margins).toEqual({
  top: 1,
  right: 2,
  bottom: 3,
  left: 4,
});
```

Mutate caller values and prove detachment. Immediately call `setCellMargins()` with `{ top: 2 }` and require editor whole-replacement snapshot `{ top: 2 }`, then clear another cell with `undefined`. Verify border/fill, transform, row/column vectors, shape identity, and non-target parts remain unchanged.

Create a margined table inside an outer transaction that throws and require full rollback. Write/reopen and require remaining margins, border/fill, geometry, and stable identity by shape id.

- [ ] **Step 2: Run typecheck and confirm the public type is red**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
```

Expected: `AddTableCellOptions` rejects `margin`.

- [ ] **Step 3: Widen only the public creation option**

In `slide.ts`, add `TextBoxMarginInput` to the existing import from `text.ts` and update:

```ts
export interface AddTableCellOptions {
  readonly border?: TableCellBorderInput;
  readonly fill?: TableCellFill;
  readonly margin?: TextBoxMarginInput;
}
```

Do not change `AddTableCell`, `AddTableCellInput`, `addTable()`, `TableCell.margins`, or `TableModel.setCellMargins()` signatures.

- [ ] **Step 4: Run model typecheck and lifecycle tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run packages/model/src/model.test.ts
```

Expected: typed public creation, detachment, immediate editor behavior, rollback, isolation, and reopen all pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

```sh
git diff --check
git add -- packages/model/src/slide.ts packages/model/src/model.test.ts
git diff --cached --check
git commit -m "feat: expose table cell margin creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

---

### Task 4: Cover the SDK lifecycle and invalid public input

**Files:**
- Modify: `packages/sdk/src/index.test.ts`

**Interfaces:**
- Consumes: Task 3 public option.
- Produces end-to-end SDK proof for native margin creation, edits, duplication, rollback, package isolation, and reopen.

- [ ] **Step 1: Add margin inputs to the existing created-table SDK lifecycle**

Type one `AddTableCellOptions` value and add cells with:

```ts
const sourceMargin = { top: 4, left: 8 };
const sourceCellOptions: AddTableCellOptions = {
  margin: sourceMargin,
  border: { kind: 'line', color: { kind: 'scheme', value: 'accent2' }, width: 2 },
  fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
};
```

Include tuple `[1, 2, 3, 4]`, scalar `0`, partial negative `{ right: -2 }`, empty object, and omitted margin. Require canonical default overlay, source detachment, existing border/fill snapshots, exact column/row geometry, and stable object identity.

- [ ] **Step 2: Exercise immediate edit, duplicate, rollback, and reopen**

After creation:

```ts
table.setCellMargins(0, 0, { bottom: 9 });
expect(table.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
const duplicate = document.duplicateSlide(0);
const duplicateTable = duplicate.shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
expect(duplicateTable).toBeInstanceOf(TableModel);
expect(duplicateTable!.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
```

Mutate the original after duplication and prove duplicate isolation. Throw inside `document.transaction()` after changing a margin and require rollback. Write/reopen both slides and compare margin matrices, text, border/fill, geometry, and non-target part hashes.

- [ ] **Step 3: Add invalid public margin values**

Build the invalid matrix explicitly:

```ts
let sdkMarginGetterCalls = 0;
const accessorNamed = { left: 1 };
Object.defineProperty(accessorNamed, 'top', {
  get() {
    sdkMarginGetterCalls += 1;
    return 2;
  },
  enumerable: true,
});
const accessorTuple = [1, 2, 3, 4];
Object.defineProperty(accessorTuple, '1', {
  get() {
    sdkMarginGetterCalls += 1;
    return 2;
  },
  enumerable: true,
});
class SdkMarginClass {
  top = 1;
}
const sparseTuple = [1, 2, 3, 4];
delete sparseTuple[2];
const invalidMargins = [
  null,
  false,
  '1',
  Number.NaN,
  Infinity,
  -Infinity,
  accessorNamed,
  accessorTuple,
  new SdkMarginClass(),
  Object.create({ top: 1 }),
  { top: 1, [Symbol('margin')]: 2 },
  Object.assign([1, 2, 3, 4], { extra: true }),
  sparseTuple,
  [1, 2, 3],
  [1, 2, 3, 4, 5],
  2_147_483_648 / 12_700 + 1,
];
```

Pass every value through `slide.addTable([[{ text: 'Invalid', options: { margin } }]])` with a narrow runtime cast. Require throws before package mutation, `sdkMarginGetterCalls === 0`, exact slide bytes, mutation journal, slide count, and shape identity.

- [ ] **Step 4: Run focused SDK and model tests**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/sdk/src/index.test.ts
```

Expected: all public lifecycle and invalid-input assertions pass.

- [ ] **Step 5: Review, commit, push, and prove synchronization**

```sh
git diff --check
git add -- packages/sdk/src/index.test.ts
git diff --cached --check
git commit -m "test: cover sdk table margin creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

---

### Task 5: Prove PptxGenJS 4.0.1 conformance and intentional differences

**Files:**
- Modify: `packages/pptxgenjs-adapter/src/index.test.ts`

**Interfaces:**
- Consumes: native public creation plus PptxGenJS public `addTable()` / `write()` output.
- Produces explicit final-direct-state equivalence for supported margins without private PptxGenJS fields.

- [ ] **Step 1: Add native and PptxGenJS tables with equivalent final margins**

Create the native matrix:

```ts
const nativeTable = nativeSlide.addTable([[
  { text: 'Default' },
  { text: 'Zero', options: { margin: 0 } },
  { text: 'One point', options: { margin: 1 } },
  { text: 'Seven point two', options: { margin: 7.2 } },
  { text: 'TRBL', options: { margin: [3.6, 7.2, 10.8, 14.4] } },
  { text: 'Negative', options: { margin: -7.2 } },
]], {
  x: inches(0.5),
  y: inches(0.5),
  columnWidths: inches(1.75),
  rowHeights: inches(1),
});
```

Create the PptxGenJS table with the equivalent public values:

```ts
const generated = new PptxGenJS();
generated.layout = 'LAYOUT_WIDE';
generated.addSlide().addTable([[
  { text: 'Default', options: {} },
  { text: 'Zero', options: { margin: 0 } },
  { text: 'One point', options: { margin: 1 } },
  { text: 'Seven point two', options: { margin: 0.1 } },
  { text: 'TRBL', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
  { text: 'Negative', options: { margin: -0.1 } },
]], {
  x: 0.5,
  y: 0.5,
  w: 10.5,
  h: 1,
  colW: 1.75,
  rowH: 1,
});
const importedTable = (await importPptxGenJS(generated)).slides[0]!.shapes.find(
  (shape): shape is TableModel => shape instanceof TableModel,
);
expect(importedTable).toBeInstanceOf(TableModel);
```

Require native/imported margin snapshot matrices, physical geometry, text, and write/reopen state to match.

- [ ] **Step 2: Assert exact direct attributes and the dual-unit difference**

Require both outputs to contain:

```ts
const directTokens = [
  'marL="91440" marR="91440" marT="45720" marB="45720"',
  'marL="0" marR="0" marT="0" marB="0"',
  'marL="12700" marR="12700" marT="12700" marB="12700"',
  'marL="91440" marR="91440" marT="91440" marB="91440"',
  'marL="182880" marR="91440" marT="45720" marB="137160"',
  'marL="-91440" marR="-91440" marT="-91440" marB="-91440"',
];
```

Add a separate native `margin: 0.1` cell and prove it writes `1270` EMU per side, while PptxGenJS `margin: 0.1` writes `91440`; document this as intentional point-vs-legacy-inch behavior. Also require a native partial named margin to retain canonical values on missing sides.

- [ ] **Step 3: Run adapter and native model suites**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run \
  packages/model/src/model.test.ts \
  packages/pptxgenjs-adapter/src/index.test.ts
```

Expected: all conformance and documented-difference assertions pass.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

```sh
git diff --check
git add -- packages/pptxgenjs-adapter/src/index.test.ts
git diff --cached --check
git commit -m "test: compare table margin creation with pptxgenjs"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

---

### Task 6: Smoke the actual Node, browser, declaration, and CLI package surfaces

**Files:**
- Modify: `scripts/smoke-npm-package.mjs`

**Interfaces:**
- Consumes: actual packed `@jiayunxie/pptx` tarball.
- Produces `tableCellMarginCreation: true` only after Node/browser/declaration creation and round-trip checks succeed.

- [ ] **Step 1: Extend Node package smoke**

Add scalar, tuple, named, empty, and margin+border+fill cells to the existing created table. Require immediate snapshots, caller detachment, canonical missing-side defaults, immediate `setCellMargins()` whole replacement, and write/reopen. Gate:

```js
const tableCellMarginCreation = tableCellBorderCreation &&
  JSON.stringify(initialCreatedMargins[0][0]) === JSON.stringify({
    top: 4,
    right: 7.2,
    bottom: 3.6,
    left: 8,
  }) &&
  JSON.stringify(initialCreatedMargins[0][1]) === JSON.stringify({
    top: 1,
    right: 2,
    bottom: 3,
    left: 4,
  }) &&
  JSON.stringify(reopenedCreatedTable.rows[0].cells[0].margins) ===
    JSON.stringify({ bottom: 9 });
```

Keep fill/border and all previous table flags in the dependency chain.

- [ ] **Step 2: Extend browser and declaration smoke**

Repeat creation/snapshot/edit/reopen assertions in the browser bundle program. In declaration smoke import `TextBoxMarginInput` and use:

```ts
const creationMargin: TextBoxMarginInput = { top: 4, left: 8 };
const creationOptions: AddTableCellOptions = {
  border: creationBorder,
  fill: cellFill,
  margin: creationMargin,
};
```

Keep `creationMargin` and `creationOptions` in the final `void` expression.

- [ ] **Step 3: Build, pack, and smoke the actual tarball**

```sh
cd packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
margin_package_dir=$(mktemp -d /tmp/pptx-table-cell-margin-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$margin_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$margin_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require `tableCellMarginCreation: true`, all previous flags true, declarations/browser true, and CLI `0.1.0`.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

```sh
cd /Users/jeremy/workspace/pptx
git diff --check
git add -- scripts/smoke-npm-package.mjs
git diff --cached --check
git commit -m "test: smoke packed table margin creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

---

### Task 7: Document the public contract and unsupported boundary

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/api/README.md`
- Modify: `docs/compatibility/pptxgenjs-baseline.md`
- Modify: `packages/pptx/README.md`

**Interfaces:**
- Consumes: tested public behavior from Tasks 1–6.
- Produces release/API/package/compatibility documentation with no stale “border/fill only” claims.

- [ ] **Step 1: Update examples and API contract**

Add `margin: { top: 4, left: 8 }` to one cell alongside border/fill. Document:

- point-only scalar/TRBL/named input;
- canonical top/bottom 3.6pt and left/right 7.2pt overlay;
- omitted/undefined/empty named byte identity;
- signed-Int32 EMU quantization and zero/negative support;
- descriptor-safe detached input;
- creation overlay versus editor whole replacement;
- marL/marR/marT/marB output and border/fill child order.

- [ ] **Step 2: Update compatibility and unsupported lists**

Change `AddTableCellOptions` examples and matrices from `{ border, fill }` to `{ border, fill, margin }`. Mark cell margin direct creation/read/edit as supported. Explain PptxGenJS `<1` inches / `>=1` points legacy behavior and exact equivalent pairs. Keep alignment, direction, fit, merge, hyperlink, rich text, table-level options, auto-page, repeated headers, and content measurement unsupported for native creation.

- [ ] **Step 3: Scan for contradictions and run typecheck**

```sh
git diff --check
rg -n -i "border/fill only|other than border/fill|options.*border.*fill|cell options" \
  CHANGELOG.md docs packages/pptx/README.md
rg -n "AddTableCellOptions.*margin|options.*margin|tableCellMarginCreation" \
  packages scripts CHANGELOG.md docs
node node_modules/typescript/bin/tsc -b --pretty false
```

Fix every current-document contradiction while leaving historical plans/specs as historical records.

- [ ] **Step 4: Review, commit, push, and prove synchronization**

```sh
git add -- \
  CHANGELOG.md \
  docs/api/README.md \
  docs/compatibility/pptxgenjs-baseline.md \
  packages/pptx/README.md
git diff --cached --check
git commit -m "docs: document table cell margin creation"
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git push ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:main
GIT_SSH_COMMAND='ssh -o BatchMode=yes -o ConnectTimeout=15 -o ServerAliveInterval=5 -o ServerAliveCountMax=2 -o UserKnownHostsFile=/tmp/codex-pptx-github-known-hosts -o StrictHostKeyChecking=yes -p 443' \
git fetch ssh://git@ssh.github.com:443/Xiejiayun/pptx.git main:refs/remotes/origin/main
git rev-list --left-right --count origin/main...HEAD
```

Expected: `0 0`.

---

### Task 8: Run full gates and real-deck QA

**Files:**
- Review every Task 1–7 path; never stage or delete `.pnpm-store/`.

**Interfaces:**
- Consumes: all implementation/tests/docs, actual packed package, repository CLI, PptxGenJS, LibreOffice, Poppler, and overflow checker.
- Produces a verified set of already-pushed commits; QA fixes, if any, receive their own review/commit/push cycle.

- [ ] **Step 1: Run full functional and performance gates**

```sh
node node_modules/typescript/bin/tsc -b --pretty false
node node_modules/vitest/vitest.mjs run --reporter=json \
  --outputFile=/tmp/pptx-table-cell-margin-vitest.json
jq '{numTotalTestSuites,numPassedTestSuites,numFailedTestSuites,numPendingTestSuites,numTotalTests,numPassedTests,numFailedTests,numPendingTests,numTodoTests,success}' \
  /tmp/pptx-table-cell-margin-vitest.json
RUN_PERF=1 node node_modules/vitest/vitest.mjs run \
  packages/testkit/src/performance.test.ts --reporter=dot
```

Require zero failed suites/tests and a passing performance gate. Do not change repository timeouts to hide host load.

- [ ] **Step 2: Rebuild and smoke the actual tarball**

```sh
cd /Users/jeremy/workspace/pptx/packages/pptx
node ../../node_modules/tsup/dist/cli-default.js
node ../../node_modules/tsup/dist/cli-default.js --config tsup.browser.config.ts
node ../../scripts/build-npm-package-types.mjs
margin_qa_package_dir=$(mktemp -d /tmp/pptx-table-cell-margin-qa-package.XXXXXX)
npm pack --ignore-scripts --pack-destination "$margin_qa_package_dir"
node ../../scripts/smoke-npm-package.mjs \
  "$margin_qa_package_dir/jiayunxie-pptx-0.1.0.tgz"
```

Require every smoke flag true, then run:

```sh
cd /Users/jeremy/workspace/pptx
node packages/pptx/dist/cli.js --json doctor
```

Require offline success and CLI `0.1.0`.

- [ ] **Step 3: Generate six real PPTX files with the SDK under test**

Create `/tmp/pptx-table-cell-margin-qa.mjs` with `apply_patch` using this complete program:

```js
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import {
  PptxDocument,
  TableModel,
  inches,
} from 'file:///Users/jeremy/workspace/pptx/packages/pptx/dist/index.js';

const require = createRequire(
  '/Users/jeremy/workspace/pptx/packages/pptxgenjs-adapter/package.json',
);
const PptxGenJS = require('pptxgenjs');
const output = '/tmp/pptx-table-cell-margin';
await mkdir(output, { recursive: true });

const textRows = [
  ['Default', 'Zero', 'Uniform'],
  ['TRBL', 'Partial', 'Negative'],
];
const tableOptions = {
  name: 'Cell margin creation QA',
  x: inches(0.75),
  y: inches(0.75),
  columnWidths: [inches(2), inches(3.5), inches(2.5)],
  rowHeights: [inches(0.8), inches(1.6)],
};

function build(rows) {
  const document = PptxDocument.create({ slideSize: 'wide' });
  const slide = document.addSlide();
  return { document, table: slide.addTable(rows, tableOptions) };
}

const strings = build(textRows);
const emptyMargin = build(textRows.map((row) => row.map((text) => ({
  text,
  options: { margin: {} },
}))));
await writeFile(output + '/string-source.pptx', await strings.document.write());
await writeFile(output + '/empty-margin-source.pptx', await emptyMargin.document.write());

const rows = [
  [
    { text: 'Default' },
    { text: 'Zero', options: {
      margin: 0,
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    } },
    { text: 'Uniform', options: {
      margin: 12,
      border: { kind: 'line', color: { kind: 'srgb', value: 'C00000' }, width: 2 },
    } },
  ],
  [
    { text: 'TRBL', options: { margin: [3.6, 7.2, 10.8, 14.4] } },
    { text: 'Partial', options: {
      margin: { top: 4, left: 18 },
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
    } },
    { text: 'Negative', options: {
      margin: { right: -3, bottom: 20 },
      border: { bottom: { kind: 'line', color: { kind: 'srgb', value: '70AD47' }, width: 3, style: 'dash' } },
    } },
  ],
];
const source = build(rows);
const sourceMargins = source.table.rows.map(({ cells }) =>
  cells.map(({ margins }) => margins));
await writeFile(output + '/margin-source.pptx', await source.document.write());

source.table.setCellMargins(1, 2, { bottom: 9 });
source.table.setCellText(1, 2, 'Edited negative');
const editedBytes = await source.document.write();
await writeFile(output + '/margin-edited.pptx', editedBytes);
const reopened = await PptxDocument.open(editedBytes);
const reopenedTable = reopened.slides[0].shapes.find(
  (shape) => shape instanceof TableModel,
);
assert.ok(reopenedTable instanceof TableModel);
assert.equal(reopenedTable.rows[1].cells[2].text, 'Edited negative');
assert.deepEqual(reopenedTable.rows[1].cells[2].margins, { bottom: 9 });
assert.deepEqual(reopenedTable.columnWidths, tableOptions.columnWidths);
assert.deepEqual(reopenedTable.rowHeights, tableOptions.rowHeights);
await writeFile(output + '/margin-reopened.pptx', await reopened.write());

const baseline = new PptxGenJS();
baseline.layout = 'LAYOUT_WIDE';
baseline.addSlide().addTable(
  [
    [
      { text: 'Default', options: {} },
      { text: 'Zero', options: {
        margin: 0,
        fill: { color: 'D9EAF7' },
      } },
      { text: 'Uniform', options: {
        margin: 12,
        border: { type: 'solid', color: 'C00000', pt: 2 },
      } },
    ],
    [
      { text: 'TRBL', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
      { text: 'Partial equivalent', options: {
        margin: [4, 7.2, 3.6, 18],
        fill: { color: baseline.SchemeColor.accent1, transparency: 25 },
      } },
      { text: 'Negative equivalent', options: {
        margin: [3.6, -3, 20, 7.2],
        border: [
          { type: 'none' },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 3 },
          { type: 'none' },
        ],
      } },
    ],
  ],
  {
    x: 0.75,
    y: 0.75,
    w: 8,
    h: 2.4,
    colW: [2, 3.5, 2.5],
    rowH: [0.8, 1.6],
  },
);
await writeFile(
  output + '/pptxgenjs-baseline.pptx',
  await baseline.write({ outputType: 'uint8array', compression: true }),
);

process.stdout.write(JSON.stringify({
  sourceMargins,
  reopenedMargins: reopenedTable.rows.map(({ cells }) =>
    cells.map(({ margins }) => margins)),
  editedText: reopenedTable.rows[1].cells[2].text,
  columnWidths: reopenedTable.columnWidths,
  rowHeights: reopenedTable.rowHeights,
}));
```

Run it and require `Edited negative`, column widths `[1828800,3200400,2286000]`, row heights `[731520,1463040]`, canonical defaults on missing creation sides, and `{ bottom: 9 }` after editor replacement/reopen.

- [ ] **Step 4: Validate packages and exact diff isolation**

Run all six exact decks:

```sh
for margin_deck in \
  string-source.pptx \
  empty-margin-source.pptx \
  margin-source.pptx \
  margin-edited.pptx \
  margin-reopened.pptx \
  pptxgenjs-baseline.pptx; do
  node packages/pptx/dist/cli.js --json package validate \
    "/tmp/pptx-table-cell-margin/$margin_deck" \
    --profile powerpoint-2010
done
```

Require zero errors/warnings. Diff:

```sh
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-margin/string-source.pptx \
  /tmp/pptx-table-cell-margin/empty-margin-source.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-margin/margin-source.pptx \
  /tmp/pptx-table-cell-margin/margin-edited.pptx
node packages/pptx/dist/cli.js --json package diff \
  /tmp/pptx-table-cell-margin/margin-edited.pptx \
  /tmp/pptx-table-cell-margin/margin-reopened.pptx
```

Require zero changed parts, only `/ppt/slides/slide1.xml`, and zero changed parts respectively.

- [ ] **Step 5: Render and inspect native/baseline decks**

Use isolated LibreOffice profiles to convert `margin-source.pptx`, `margin-edited.pptx`, and `pptxgenjs-baseline.pptx` to PDF; rasterize every page with Poppler at 180 DPI. Run both overflow checks:

```sh
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-margin/margin-edited.pptx
/Users/jeremy/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
  /Users/jeremy/.codex/plugins/cache/openai-primary-runtime/presentations/26.727.11326/skills/presentations/container_tools/slides_test.py \
  /tmp/pptx-table-cell-margin/pptxgenjs-baseline.pptx
```

Inspect every PNG at original detail. Require visibly different zero/uniform/TRBL/partial/negative margins, preserved border/fill, unequal rows/columns, all text, and no repair, clipping, unexpected wrap, overlap, blur, missing side/cell, or off-slide content.

- [ ] **Step 6: Final static review and synchronization proof**

```sh
git diff --check
git status --short
git rev-list --left-right --count origin/main...HEAD
```

Review descriptor reads, canonical default bytes, point quantization, named overlay, editor whole replacement, public declarations, smoke output, PptxGenJS differences, docs boundaries, and invalid-input no-mutation assertions. Expected final output is `0 0` and only `?? .pnpm-store/`. If QA reveals a defect, fix only that defect, run its focused/full gates, review, commit, push, fetch, and repeat this step.

---

After synchronization, design the next independently reviewable cell creation option without asking for routine decisions.
