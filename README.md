# xtl-py

A Python reference implementation of **XTL (Excel Template Language) 0.1** —
a host-language–agnostic templating language that turns an `.xlsx` template
plus an `.xlsx` data workbook into one or more rendered output workbooks.

`xtl-py` is the second implementation of the language, built alongside the
TypeScript reference impl as portability validation: if both implementations
produce identical output for the same conformance corpus, the spec is real.

- **Spec & TS reference**: <https://github.com/jinyoung4478/xl3>
  (XTL spec **0.1.0** released 2026-05-08).
- **Conformance status**: **91 / 91 stage-1 fixtures passing (100%)** —
  5 stage-2 fixtures skipped (canonical OOXML comparison out of scope for v0.1).
  The 5 spec/impl ambiguities found while building this port were filed as
  [xl3 issue #1](https://github.com/jinyoung4478/xl3/issues/1) and resolved
  upstream the same day; see `PORTING_NOTES.md`.

## Install

```bash
pip install xtl-py
```

> ⚠️ The PyPI distribution name is **`xtl-py`** but the import name is
> **`xl3`** (matching the TS package on npm).
>
> ```python
> from xl3 import convert  # NOT `import xtl_py`
> ```

Requires Python ≥ 3.11.

## Quick start

```python
from xl3 import convert

with open("template.xlsx", "rb") as f:
    template = f.read()
with open("data.xlsx", "rb") as f:
    data = f.read()

output_files = convert(template, data)

for f in output_files:
    with open(f.filename, "wb") as out:
        out.write(f.data)
    print("wrote", f.filename)
```

### Runtime inputs (ADR-0010)

```python
from xl3 import convert, ConvertOptions

output = convert(
    template,
    data,
    ConvertOptions(inputs={"month": "2026-05", "region": "Seoul"}),
)
```

### Inspecting a template

```python
from xl3 import preview, read_template_inputs

# Lightweight: returns parsed file/sheet/row counts + warnings without rendering.
result = preview(template, data)

# Just the input declarations (for building a host UI).
specs = read_template_inputs(template)
```

### Structured errors

```python
from xl3 import convert, is_xtl_error

try:
    convert(template, data)
except Exception as e:
    if is_xtl_error(e):
        print(e.code, e)  # e.g. "xl3/source/sheet-missing", message
    else:
        raise
```

Every spec-defined error carries a stable `xl3/<category>/<id>` code per
ADR-0015. The English message is the conformance contract; hosts should
dispatch on `code`.

## Conformance runner

The package ships a CLI runner that implements `conformance/runner-protocol.md`:

```bash
# Full stage-1 run against an xl3 fixture directory
python -m xl3.runner --fixture-dir /path/to/xl3/conformance/fixtures

# JSON report
python -m xl3.runner --report json

# Filter
python -m xl3.runner --filter substitution
python -m xl3.runner --id-prefix 050
```

Output sample:

```
xl3-py 0.1.0a2 — XTL 0.1 (stage 1)
  pass   001-bracket-substitution
  pass   002-if-function
  ...
summary: 91/91 passed, 0 failed, 5 skipped
```

## What is supported

| Surface | Status |
|---|---|
| Bracket substitution `{{ [Col] }}` | ✅ |
| `IF` / `IFEMPTY` / `ROUND` / `ABS` / `TEXT` / `TODAY` / `ROW` | ✅ |
| Aggregates `SUM` / `COUNT` / `AVERAGE` / `MIN` / `MAX` | ✅ |
| `XLOOKUP` (3-arg + 4-arg) | ✅ |
| Directives `@filter` / `@sort` / `@top` / `@repeat right` / `@source` / `@join` | ✅ |
| `__config__` / `__inputs__` / `__sources__` / `__lists__` reserved sheets | ✅ |
| ADR-0007 / 0008 / 0009 / 0017 value model | ✅ (82 unit tests pinning the contract) |
| ADR-0002 filename sanitization | ✅ |
| ADR-0003 numFmt-driven coercion | ✅ |
| ADR-0010 runtime inputs (text / number / date / select) | ✅ |
| ADR-0012 multi-source data model | ✅ |
| ADR-0013 XLOOKUP cross-source | ✅ |
| ADR-0014 single inner `@join` | ✅ |
| ADR-0016 file/sheet group splitting (first-seen order) | ✅ |
| Stage-2 canonical OOXML comparison | ❌ (out of scope for v0.1; deferred) |

## Architecture

Pure-Python, sync API, single dependency on
[`openpyxl`](https://pypi.org/project/openpyxl/) for Excel I/O.

```
src/xl3/
├── __init__.py        # public API
├── errors.py          # XtlError + ADR-0015 stable code catalog
├── value_model.py     # is_empty / is_truthy / canonical_string /
│                      # canonical_number / compare_values
├── expression.py      # cell-template lexer + recursive-descent parser
├── evaluator.py       # AST eval + ROW / aggregates / XLOOKUP
├── functions.py       # IF / IFEMPTY / ROUND / ABS / TEXT / TODAY
├── directives.py      # @filter / @sort / @top / @repeat / @source / @join
│                      # + row-set transform pipeline
├── inputs.py          # ADR-0010 input resolution
├── filename.py        # ADR-0002 sanitization
├── parser.py          # template workbook → block plan
├── reader.py          # source workbook reader (multi-source)
├── renderer.py        # block-based renderer with file/sheet groups
├── pipeline.py        # convert / preview / read_template_inputs
└── runner/            # conformance runner CLI
```

## Status

**Pre-1.0 / alpha.** API surface mirrors the TS reference. Breaking changes
are possible until the spec freezes at XTL 1.0; see
[`spec/STABILITY.md`](https://github.com/jinyoung4478/xl3/blob/main/spec/STABILITY.md)
in the spec repo.

## License

MIT — see the spec repo's `LICENSE`.
