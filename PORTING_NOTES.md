# Porting Notes — XTL TS → Python

A running log of points where the spec is genuinely underspecified,
internally inconsistent, or where the Python port had to make a non-obvious
call. Format per entry:

- **Where**: spec section / ADR / fixture number
- **Question**: the ambiguity in one sentence
- **TS impl behavior**: what the JS reference implementation does today
- **Other reasonable interpretations**: alternatives we considered
- **Our choice**: what the Python port does, and why
- **Severity**: spec/impl/test (does it block conformance? does it block
  cross-impl portability?)

When this file accumulates 5+ entries, batch them into an issue against
`xl3` and propose ADR amendments.

---

## #1. ECMA-262 scientific-notation cutoff is `1e-6`, not `1e-4`

**Where**: `spec/language.md` "Canonical String Form"; `spec/decisions/0009-comparison-and-string-coercion.md` §"Canonical string form".

**Question**: ADR-0009 says the canonical-string form of a finite number uses
"no scientific notation for magnitudes between `1e-4` and `1e21`," and claims
this matches `Number.prototype.toString`. But ECMA-262 §6.1.6.1.13 actually
uses a **`-6`** cutoff, not `-4`. So `(0.00005).toString()` is `"0.00005"` in
JS — decimal, not scientific. The spec text and the cited authority disagree.

**TS impl behavior**: uses the host `Number.prototype.toString`, so follows
ECMA-262 (`-6` cutoff). The fixture corpus was authored against this
behavior.

**Other reasonable interpretations**:
- Take the spec text literally: scientific for `0 < |x| < 1e-4`. Diverges
  from JS impl + corpus.
- Take ECMA-262 literally: `-6` cutoff. Matches impl + corpus, contradicts
  spec text.

**Our choice**: ECMA-262 cutoff (`-6`). Implementing the literal spec
text would fail any fixture exercising values in `[1e-6, 1e-4)`. Our
`canonical_number` re-implements ECMA-262 §6.1.6.1.13 directly so the
behavior is JS-compatible regardless of host language quirks.

**Severity**: spec — propose an ADR amendment that replaces "1e-4" with
"1e-6" and either drops the redundant `Number.prototype.toString` cite or
keeps it now that the text matches.

---

## #2. ECMA `String.prototype.trim` historically includes U+FEFF; ADR-0007 excludes it

**Where**: `spec/decisions/0007-empty-value-definition.md`.

**Question**: ADR-0007 says whitespace "matches the set recognized by
ECMAScript `String.prototype.trim` — equivalent to the Unicode-mode `\s`
character class," but then explicitly excludes U+FEFF (zero-width no-break
space / BOM) and other zero-width characters. ECMAScript's `WhiteSpace`
production has historically *included* U+FEFF, and most engines implement
`trim()` accordingly. So a string of bare U+FEFF is empty per native JS
trim but **non-empty** per ADR-0007.

**TS impl behavior**: uses native `String.prototype.trim`, which strips
U+FEFF on V8/JSC/SM. So a bare-U+FEFF source cell is empty per the impl,
non-empty per the spec.

**Other reasonable interpretations**:
- Follow Unicode `White_Space` property strictly (what the ADR text says).
- Follow ECMA-262 `WhiteSpace` production (what the impl actually does).

**Our choice**: Unicode `White_Space` (Python `str.isspace()`), per the
ADR's normative bullet. No fixture currently asserts this edge case, so the
divergence is silent for the bootstrap corpus, but we follow the spec.

**Severity**: impl — TS impl deviates from spec on a technicality. Worth a
fixture (`empty-zwnbsp-not-whitespace`) to pin behavior either way.

---

## #3. Python `repr(float)` exponent padding differs from ECMA `Number.prototype.toString`

**Where**: implementation detail; affects every numeric `&` concat and every
single-expression cell coercion fallback.

**Question**: Python `repr(1e-7)` is `"1e-07"` (two-digit padded exponent);
JS `(1e-7).toString()` is `"1e-7"`. Python `repr(-0.0)` is `"-0.0"`; JS
`(-0).toString()` is `"0"`.

**TS impl behavior**: matches ECMA-262 by virtue of being JS.

**Our choice**: re-implement ECMA-262 §6.1.6.1.13 directly in
`canonical_number`. Don't trust `repr(float)` to format identically — we
extract digits + decimal exponent from `repr` and re-render using the
ECMA format selection rules (decimal in [1e-6, 1e21), scientific outside).
Negative zero returns `"0"`.

**Severity**: impl — just a thing the port has to handle.

---

## #4. Excel serial dates are timezone-naive; ADR-0017 mandates UTC

**Where**: ADR-0017 §"Timezone (normative)".

**Question**: ADR-0017 says Date components MUST be read in UTC.
ExcelJS exposes Excel's timezone-naive serial dates as `Date` objects
*anchored at UTC midnight*, and the impl uses `getUTCFullYear` etc.
openpyxl returns Python `datetime.datetime` objects with `tzinfo=None`
(naive); calling `.year`/`.month`/`.date()` on them yields the serial's
naked components, which is what ADR-0017 wants — but only if you don't
"localize" them first.

**TS impl behavior**: explicit `getUTC*` calls.

**Our choice**: never call `.astimezone()` on a naive datetime; treat
naive openpyxl datetimes as already-UTC. `canonical_date` strips
`tzinfo` defensively if present.

**Severity**: impl — easy to get wrong subtly. Worth a Stage 1 timezone
matrix fixture run (CI matrix per `STABILITY.md`).

---

## #5. TS IF-condition normalizer recognizes `==` but NOT `=` — spec uses `=`

**Where**: `xl3/src/normalizer.ts` lines ~215-225; fixture
`048-if-and-comparison-boundaries`.

**Question**: `language.md` §"Comparison Operators" lists `=` (single equal)
as the equality operator. The TS normalizer's IF-condition op list is
`[['!=', 'ne'], ['>=', 'ge'], ['<=', 'le'], ['==', 'eq'], ['>', 'gt'],
['<', 'lt']]` — it has `==` but NOT `=`. So
`IF([Amount] = 0, "zero", "non-zero")` is normalized as a function call
where the condition is the *raw string* `"[Amount] = 0"`. At eval time
it falls through to "bare string literal" which is non-empty → truthy →
IF always takes the THEN branch regardless of Amount.

Fixture 048 was authored against this buggy behavior. Its
`expected.xlsx` has `D3='zero'` for an Amount=1 row, even though
spec-correct evaluation of `IF(1 = 0, "zero", "non-zero")` is
`"non-zero"`.

**TS impl behavior**: bug — `=` not recognized in IF condition; treated
as literal string (always truthy).

**Other reasonable interpretations**: this is unambiguously a TS impl
bug under the spec.

**Our choice**: follow the spec — `=` is the equality operator
everywhere, including IF conditions. We FAIL fixture 048.

**Severity**: impl + test — `xl3/src/normalizer.ts` should add `=` to
the ops list before `==`; and fixture 048's `expected.xlsx` should be
re-authored to reflect spec-correct evaluation. Note that `@filter [col]
= value` works correctly in TS (different parser). Other fixtures using
`=` (063, 064, 079, 080, 081, 088, 092) are inside @filter / @join
clauses that TS parses with a separate code path that does handle `=`.
Fixture 048 is the only one currently affected.

## #6. (placeholder — append future ambiguities below)
