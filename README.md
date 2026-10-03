# gba-mep-docx-toc

A [DeepSeek Harness](https://github.com/deepseek-ai/deepseek-harness) (`dsh`) bundle that registers one skill: **insert a real Word table of contents into an existing `.docx`**.

The skill writes a genuine `TOC` field with `_Toc` bookmarks straight into `word/document.xml` — graded indent per heading level, dot leaders, right-aligned page numbers — using only the Python standard library. It does not need Microsoft Word, LibreOffice, COM, or any third-party package.

## Install

```sh
dsh plugin --profile web add gba-mep-docx-toc
```

This repository is also installable straight from source through pnpm's GitHub spec, and it is listed in the [awesome-dsh-plugin](https://awesome-dsh-plugin.com) catalog, so [dsh-market](https://dshmarket.com) can install it with one click.

After installing, restart `dsh web` once so the bundle layer is composed.

## What it does

Given a `.docx` whose headings use the `Heading 1`–`Heading 9` styles (or carry `w:outlineLvl`):

1. finds the headings in document order,
2. inserts `_Toc…` bookmarks so every entry is clickable,
3. builds the `TOC \o "1-3" \h \z \u` field with one entry per heading — hyperlink anchor → title text → right-aligned tab with dot leader → `PAGEREF` field,
4. indents level 2 by `100` and level 3 by `200` character units,
5. rewrites only `word/document.xml` and copies every other part of the package byte for byte,
6. verifies the result with `verify` (entry count = bookmark count = `PAGEREF` count).

## Requirements

- A DSH build that exposes the `skills` service — the entry imports nothing from the harness, so it works on any host publishing that contract (checked against DSH `0.2.0-rc.2`).
- Python 3.9+ on the machine that runs the script. No packages to install.
- Node.js only because it is a dsh plugin; the entry itself uses `node:fs`, `node:path` and `node:url`.

## Usage

```sh
python scripts/docx_toc.py build 输入.docx -o 输出.docx --title "目　录" --page-break both
python scripts/docx_toc.py verify 输出.docx
python scripts/docx_toc.py selftest
```

In a session you can also just describe the task — the skill's `description` is what the agent routes on.

## Boundaries

- **Page numbers are placeholders.** Python does not paginate, so real numbers appear only after Word updates the field (`Ctrl+A`, then `F9`, choose *Update page numbers only*). Until then every entry reads `1`.
- **No section-based page numbering.** With sections that restart numbering (roman front matter, arabic body), the `PAGEREF` result depends on the Word section setup and is not guaranteed.
- **It does not edit an existing TOC.** A document that already carries a `TOC` field gets a second one; delete the old one first.
- **Paginating is not guaranteed to match Word.** Rendering checks used LibreOffice, whose page breaks differ from Microsoft Word's.
- **Headings must use heading styles.** Manually bolded and enlarged text is invisible to the script.

## Layout

```
lib/index.js            bundle entry — registers the skill via ctx.skills.register()
cordis.patch.yml        loader patch: one insert row pointing at this package by name
SKILL.md                skill body (frontmatter: name / description / whenToUse)
scripts/docx_toc.py     the implementation (stdlib only)
references/OOXML-TOC.md notes on the OOXML table-of-contents field
tests/                  offline tests: entry behaviour and installability contract
screenshots.json        screenshots shown by plugin storefronts
```

## Tests

```sh
node --test tests/
```

`tests/entry.test.mjs` runs `apply()` against a minimal context and asserts the registered fields, the `resourceBase`, and that the referenced files ship with the package — including a CRLF regression for the frontmatter parser. `tests/bundle-contract.test.mjs` asserts the installability contract: `dsh.bundle` declared, patch row naming the package, `files` whitelist complete, entry importing only `node:` builtins.

## License

MIT — see [LICENSE.md](LICENSE.md).
