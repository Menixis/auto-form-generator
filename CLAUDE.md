# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A small Thai-language tool for ปตท. (PTT) staff that fills five standard Word templates (`templates/form1.docx`–`form5.docx`) from a PO number + a few manual fields, auto-filling vendor / job title from an Excel sheet of PO data. Two UIs share one logic module.

Parent repo conventions and the project's place in the broader tree are documented in `..\..\CLAUDE.md` — read that first; this file only covers what's specific to `auto-form-generator/`.

## The split: two UIs, one logic core

```
form_logic.py     ← pure logic, no UI deps (only python-docx)
   ↑          ↑
   │          │
auto_form_generator.py   streamlit_app.py
(tkinter desktop)        (Streamlit web — also the HF Spaces entry)
```

**Invariant:** `form_logic.py` must never import `tkinter` or `streamlit`. Both UIs depend on it. Anything UI-agnostic — form metadata (`FORM_TYPES`), placeholder constants, the `generate_form()` entry point, the `replace_text()` helper — lives there. UI-specific concerns (file dialogs, session state, progress spinners, Excel-source toggling) stay in the respective UI module.

`generate_form(output_path=...)` accepts either a filesystem path (desktop) or a `BytesIO` (web). `python-docx` `Document.save()` handles both — preserve that polymorphism.

## Excel lookup has two completely separate code paths

This is the non-obvious part. The desktop and web versions do **not** share lookup code:

- **Desktop** (`auto_form_generator.py` → `excel_lookup.py`): reads `config.ini` for `file_path` / `sheet_name` / `po_column`, opens the file from disk (OneDrive-synced or network share). `excel_lookup` is a soft import — if it or `openpyxl` is missing, the "ดึงข้อมูล" button disables itself and the rest of the app still works.
- **Web** (`streamlit_app.py` → `lookup_po_in_bytes()`): no `config.ini`, no filesystem. Excel comes in as uploaded bytes (or the bundled `sample_po_data.xlsx`) and is parsed from a `BytesIO`. `po_column` / `sheet_name` come from sidebar inputs.

If you find yourself adding a feature that touches PO lookup, decide deliberately whether it belongs in one path, both, or pulled up into a shared helper. Right now the duplication is intentional — the two environments have different I/O models.

## How the docx replacement works (and why it's fragile)

`form_logic.replace_text()` walks every paragraph (including those nested in tables) and tries two strategies in order:

1. **Single-run match**: if `old_text` fits entirely inside one `run`, replace it there and preserve that run's formatting.
2. **Cross-run fallback**: if the placeholder is split across multiple runs (common when Word re-segments after edits), it rewrites `paragraph.runs[0].text` with the full replacement and blanks out the rest of the runs. This **loses the formatting of runs 2..N** in that paragraph.

When editing templates, prefer keeping each placeholder inside a single run so strategy 1 fires. If you change template text in Word, re-open and check the placeholder isn't accidentally split (a stray bold/italic toggle mid-placeholder will do it).

## Placeholder schemes

All five templates now share the named placeholders `Vendor`, `ชื่องาน`, `PO_No` (constants `PLACEHOLDER_VENDOR`, `PLACEHOLDER_WORK`, `PLACEHOLDER_PO`). The shared block at the end of `generate_form()` fills them via straight find-and-replace regardless of `form_id`.

Form 1 additionally has named placeholders for the document number (`หมายเลขเอกสาร`), document type (`แจ้งให้เริ่มเข้าทำงาน / เริ่มส่งมอบสินค้า`), and signer (`ผู้มีอำนาจอนุมัติ หรือ ประธานกรรมการตรวจรับ`).

Forms 2–5 use **layout-based placeholders** for those same three: the doc-number slot is `                   /` (19 spaces + slash), and the signer line is positioned by `format_signer_line_f2_5()` — exactly 67 leading spaces and a 44-wide centered parenthesised field. Don't "clean up" those magic numbers without re-checking the rendered output in Word. There is no `doc_type` slot in forms 2–5; the radio is disabled for those forms because none of form 1's two options match their subject lines.

Templates 2–5 were edited from their original PTT-issued state to insert the shared placeholders (`Vendor`, `ชื่องาน`, `PO_No`) at semantically appropriate slots — see `templates_backup/` for the originals. The `PO_No` placeholder appears only in the `ใบสั่งซื้อ/จ้าง เลขที่` reference line (not the `สัญญา` alternative, which carries a contract number that is generally distinct from the PO).

When adding a new form, prefer the shared named placeholders over inventing new ones, and reserve layout-based placeholders for cases where the official template's spacing is legally meaningful.

## Running

Beyond what's in `README.md`:

```powershell
# Smoke-test Excel config without launching the GUI:
python excel_lookup.py --list           # list sheets/columns in the configured file
python excel_lookup.py 4500987654       # try a real lookup

# Build/run the HF Spaces container locally:
docker build -t auto-form-generator .
docker run --rm -p 7860:7860 auto-form-generator
```

There are no tests. Validation is: launch one of the UIs, run a real PO through it, open the resulting `.docx` in Word, check formatting survived.

## Deployment shape

`Dockerfile` + the YAML frontmatter in `README.md` (`sdk: docker`, `app_port: 7860`) target Hugging Face Spaces. The `STREAMLIT_*` env vars in the Dockerfile (port 7860, headless, no usage stats, non-root user 1000) are HF Spaces requirements — don't strip them. The `.streamlit/config.toml` referenced in the README structure is the theme/toolbar config (it isn't present in the tree, so theme falls back to Streamlit defaults).

`config.ini` is **not** shipped (only `config.ini.template`) and is only used by the desktop path — the web/Docker path never reads it. Keep it that way; the container shouldn't depend on per-user filesystem config.
