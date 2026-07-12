---
name: excel-incell-translate
description: |-
  Translate Excel (.xlsx) spreadsheets by embedding the translation INSIDE the
  same source cell (original text on top line, translation below, separated by a
  newline), without modifying the original text and without creating new columns
  or new files. Handles bidirectional Chinese-to-English and
  English-to-Chinese in one pass, skips formulas/numbers, and
  preserves all source content. Use this skill when a user asks to translate an
  Excel file "in-place", "add translation to each cell", "keep the source
  text", or complains that translations landed in a separate/right-hand column
  instead of inside the cell.
agent_created: true
---

# Excel In-Cell Bilingual Translate

## Overview

Translate every text cell of an `.xlsx` workbook by appending the translation to
the SAME cell that holds the source text. The cell becomes two lines:
`original text` on line 1, `translation` on line 2 (joined by `\n`, with
`wrap_text` enabled). The original text is never deleted or altered — only
appended. No new columns, no new sheets, no new files.

Direction is per-cell and automatic:
- Cell contains CJK (Chinese)  → append an English translation.
- Cell contains Latin letters only  → append a Chinese translation.
- Cell is a formula, a number, blank, or pure punctuation → leave untouched.
- Cell is an untranslatable token (test-case ID, abbreviation, carrier/brand
  name, person name) → leave untouched ("don't write nonsense").
- Cell is a verdict/status token (PASS, FAIL, N/A, OK, NG, …) → leave
  untouched. These are test conclusions whose meaning is universal, and the
  user requires they NOT be modified. The script enforces this via STATUS_SKIP
  even if such a token is present in the dict.

## When To Use

- "Translate this Excel file / these sheets" where the user expects the result
  inside the original file.
- "Add the Chinese/English translation to each cell" / "keep the source
  description, just add the translation".
- A prior run placed translations in a separate right-hand column ("后面的表格")
  and the user wants them moved back INTO the source cells.
- Bidirectional glossaries: Chinese sentences → English, English sentences →
  Chinese, in one pass.

Do NOT use this skill to produce a brand-new translated workbook in a
different file unless the user explicitly wants a separate output file.

## Core Workflow

### 1. Backup, then load twice

Always keep a pristine copy. Load the workbook in two modes:
- `wb = load_workbook(path)`            # keeps formatting/styles
- `wb_data = load_workbook(path, data_only=True)`  # reads cached values

Read the source value from `wb_data` (so formula cells return their cached
value/number, not the formula string) and write the translation into `wb`.

### 2. Build the translation dictionary (domain-specific, EXTERNAL)

Extract every unique text string from `wb_data`, classify each as CJK /
Latin / other. Translate and write ONE external JSON dict
`{"source": "translation", ...}` that the embedder loads via `--dict`.
The dict is shared by BOTH body cells and headers — include the header
strings (e.g. `"Call": "通话"`) in the same file.
- A small set of phrase tables (`PHRASES = {"Pass":"通过", ...}`) plus
  manual translation of the high-frequency, domain-specific sentences gives the
  best quality. Direction is encoded by the entry itself.
- Keep IDs, abbreviations, brand/carrier/person names OUT of the dict (or map
  them to themselves) so they are skipped.
- Never fabricate a translation for an unknown token — skip it.
- The script bundles NO project data; the dict is supplied per project. See
  `references/translation_method.md` for the extraction/classification pattern.

### 3. Embed in-cell (the critical step)

**Idempotency first (best practice #1).** Before touching a cell, call
`is_already_translated(value)`: if the value already has 2+ lines whose
first and last lines are of OPPOSITE language (e.g. `PASS`⏎`通过`),
treat it as already translated and SKIP it. This makes the skill safe to
re-run on a previously translated file — it only fills in cells that still
lack a translation, never re-translates or stacks translations.

For each remaining source cell `(r, c)`:
- Skip if it is a `MergedCell` that is not the top-left of its merge
  (its value lives in the top-left; writing there raises
  `AttributeError: 'MergedCell' object attribute 'value' is read-only`).
- Look up the (stripped) text in the external dict.
- If a translation exists and differs from the source:
  `cell.value = f"{source}\n{translation}"`
  then set `wrap_text`:
  ```python
  from openpyxl.styles import Alignment
  al = cell.alignment
  cell.alignment = Alignment(horizontal=al.horizontal,
                             vertical=al.vertical, wrap_text=True)
  ```
  This preserves horizontal/vertical alignment and only adds line-wrapping.
- If no translation (ID / number / formula / blank / unknown token) →
  do nothing.

Headers (row 1) are cells too: translate them from the SAME external
dict (headers are included in the dict built in step 2), using the same
`source\ntranslation` format. No separate hardcoded header map.

### 4. Recover from a "wrong-column" prior run (explicit, safe)

If a previous run mistakenly wrote translations into a separate right-hand block
of columns, do NOT re-translate. Recover with an EXPLICIT, backup-gated
command (best practice #5 — never guess column doubling from `mm == 2*bm`,
which can delete real data when the source sheet itself has an even column
count and no backup is present):

```bash
python scripts/embed_incell.py workbook.xlsx --recover --backup clean.xlsx
```

- The script requires `--backup` in recover mode and refuses to run without
  it. `bm` (true source column count) is taken from the backup.
- For each body cell `(r, c)` with row greater than 1: the clean translation
  already sits in column `bm + c` at the same row `r`. Reuse it:
  `cell(r,c).value = f"{source}\n{wb.cell(r, bm+c).value}"`.
- Header row (r == 1) right-column values are DIRTY (`"Header |译文"`);
  do not reuse them. Translate headers from the `--dict` instead.
- After embedding, delete the right-hand block:
  `ws.delete_cols(bm + 1, current_max - bm)`.
- Formula/number cells still have empty right-column values → skipped,
  source left intact.

### 5. Preserve formulas across save

`openpyxl` DROPS cached formula values when it saves (the formula string
is kept, but its last-computed value becomes `None` until Excel recalculates).
Fix by forcing recalculation on open:
```python
try:
    wb.calculation.fullCalcOnLoad = True
except Exception:
    pass
wb.save(path)
```
Verify against the clean backup by comparing FORMULA STRINGS and non-formula
STATIC values — NOT cached values (those legitimately become None and are
restored when the user opens the file in Excel).

### 6. Verify before reporting done

- Every sheet's column count equals the backup's (`delete_cols` worked).
- Every translated cell's value STARTS WITH the original text (proves the
  source was appended-to, never overwritten).
- Formula strings match the backup exactly (0 differences).
- Every sheet that contained text has at least one embedded translation.

## Critical Pitfalls (read before coding)

1. **Never `insert_cols` to make room.** It silently corrupts/unloads
   source cells (numbers and dates turning to `None`). Always write into the
   EXISTING source cell; only `delete_cols` for cleanup.
2. **MergedCell is read-only.** Check `isinstance(cell, MergedCell)`
   before writing; skip non-top-left merge cells.
3. **Formula cached values are lost on save** → set
   `fullCalcOnLoad = True` and verify with formula-string comparison, not
   cached-value comparison.
4. **Don't translate formulas/numbers.** Reading via `data_only=True`
   returns a number for a formula cell; skip those (no text to translate).
5. **Don't fabricate.** Unknown tokens (IDs, names, abbreviations) are
   skipped, never guessed.
6. **In-cell = newline, not a parallel column.** The whole point of this
   skill is that the translation lives in the same cell as the source, so the
   file structure (column count, sheet layout) is unchanged.
7. **Idempotency: detect already-translated cells, never re-translate.**
   A re-run must skip cells already in `source\ntranslation` form (first and
   last lines of opposite language). Without this, re-running stacks or
   scrambles translations. `is_already_translated()` encodes the heuristic.
8. **No hardcoded project data in the script.** Translation dicts AND header
   maps are supplied externally via `--dict`. A script that bundles one
   project's table headers mis-applies them to unrelated workbooks and
   contradicts this skill's own "no bundled dict" rule.
9. **Recover mode is explicit and backup-gated.** Never auto-detect a
   "doubled column" state from `mm == 2*bm` — when the source sheet itself
   has an even column count and no backup is given, that guess wipes real
   columns. Require `--recover --backup`.
10. **Don't modify verdict/status characters (PASS/FAIL/…).** Test
   conclusions (PASS, FAIL, N/A, OK, NG, …) are universal and often parsed
   by automation; appending a translation harms them. The script hard-skips
   these via `STATUS_SKIP` regardless of the dict. Do NOT put them in the
   dict as translatable, and the `is_status_token()` check wins even if you do.

## Reusable Script

`scripts/embed_incell.py` is a parameterized embedder. It embeds
translations in-cell from an external dict (`--dict`), is idempotent (re-runs
skip already-translated cells), and can safely recover a prior wrong-column
run when given an explicit `--backup`. Run:

```bash
# 翻译模式（表头与正文共用同一外部字典）
python scripts/embed_incell.py workbook.xlsx --dict dict.json [--backup clean.xlsx]

# 恢复模式：从右侧错位列回收译文（必须带 --backup）
python scripts/embed_incell.py workbook.xlsx --recover --backup clean.xlsx
```

The domain translation dict is intentionally NOT bundled — supply it per
project (extract unique strings → classify → translate). See
`references/translation_method.md` for the extraction/classification pattern.
