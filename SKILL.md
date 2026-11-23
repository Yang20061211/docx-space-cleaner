---
name: docx-space-cleaner
description: Remove unwanted spaces after bullet markers, between Chinese text and Latin letters or digits, between digits and Latin words, and around slash or bracket boundaries in .docx files, while preserving spaces inside ordinary English phrases and code. Use for Chinese Word document whitespace cleanup.
---

# DOCX Space Cleaner

Clean Chinese Word documents that should not use visual gaps between Chinese and English or numbers.

## Required behavior

- Remove spaces after literal `•` markers, such as `• 创建 n 个` to `•创建n个`.
- Remove spaces at boundaries between Chinese characters and Latin letters or digits, in either direction.
- Remove spaces between digits and Latin words, such as `2.4 Kruskal` to `2.4Kruskal`.
- Remove spaces around `/`, such as `30 / 435 / 870` to `30/435/870`.
- Remove spaces between Chinese and an opening bracket or parenthesis, and between a closing bracket or parenthesis and Chinese.
- Disable Word's paragraph-level automatic spacing between East Asian and Latin text and between East Asian text and numbers.
- Preserve spaces inside English phrases, commands, code signatures, URLs, identifiers, and data. For example, keep `int n`, `already connected`, and `std::map`.
- Preserve text, tables, page structure, fonts, and all other formatting unrelated to this cleanup.
- Create a new `.docx` by default. Do not overwrite the source unless explicitly requested.

Do not use a blanket "remove every space" operation.

## Workflow

1. Use the `documents` skill for DOCX handling and render verification. Resolve the bundled Python runtime through workspace dependencies when available.
2. Choose a new output path. A useful name is `<source-stem>_去多余空格.docx`.
3. Run the bundled cleaning script:

   ```bash
   python scripts/clean_docx_spacing.py INPUT.docx OUTPUT.docx
   ```

4. Verify the observable result:

   - no `•[whitespace]` sequences remain;
   - no Chinese, digit, Latin, slash, or bracket space boundaries named above remain;
   - English phrases and code keep their internal spaces;
   - paragraph and table counts are unchanged;
   - all visible text differences are only the intended deleted whitespace.
5. Render the output and inspect every page for reflow, clipping, overlap, or broken tables. If the local renderer lacks the source document's Chinese font, diagnose that as a renderer limitation and do not alter the document's font settings.

## Scope

This skill handles text already inside `.docx` paragraphs, tables, headers, and footers. It does not convert `.doc`, `.wps`, PDF, scanned, or image-only documents.

Report the number of removed bullet spaces and inline spaces in the final response.
