# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build

**Target platform**: Windows only (Win32 API, RichEdit control). The build produces `WhatsUp.exe`.

```bash
# MSVC (Visual Studio 2022)
cmake -B build -G "Visual Studio 17 2022"
cmake --build build --config Release

# MinGW-w64
cmake -B build -G "MinGW Makefiles"
cmake --build build
```

First build auto-fetches **zlib** via CMake FetchContent (internet required). All `*.cpp` files under `src/` are picked up by `GLOB_RECURSE` — no CMakeLists.txt edit needed when adding new source files.

There are no automated tests. Verification is done by building and running on Windows.

## Architecture

```
File → FormatHandler → cdm::Document → CdmNormalizer → CdmLoader → RichEdit
```

### CDM (Canonical Document Model) — the core abstraction

All format handlers convert their native format into a `cdm::Document` tree defined in `src/cdm/document_model.hpp`. This is a rich, lossless intermediate representation. Understanding this model is essential before touching any format handler or renderer.

**Key types** (all in `namespace cdm`):

- `Document` — root; contains `sections`, `styles`, `resources`, `comments`, `notes`, `bookmarks`
- `Section` — page settings + `std::vector<BlockPtr>`
- `Block` — `std::variant<Paragraph, Heading, CodeBlock, Table, BlockQuote, ListBlock, HorizontalRule, RawBlock, SectionBreakBlock, CustomBlock>`
- `Inline` — `std::variant<Text, Tab, Break, InlineCode, Hyperlink, Image, Span, Emphasis, Strong, Underline, …>`
- `TextStyle` — all `std::optional`: `fontFamily`, `fontSize`, `color`, `bold`, `italic`, `underline`, `strike`, `script`, `charset`, …
- `ParagraphStyle` — alignment, indents, spacing, line-break mode
- `StyleDefinition` — named style with `TextStyle` + `ParagraphStyle`; nodes reference these via `styleRef` (a `std::string` ID)

**Critical design rule**: `directStyle` is the inline/direct formatting; `styleRef` is a reference to a named `StyleDefinition` in `doc.styles`. The Normalization stage is responsible for resolving `styleRef` → merging into `directStyle`, and for promoting paragraphs whose `styleRef` matches "Heading1"–"Heading6" into `Heading` blocks. **This resolve+promote step is not yet implemented** and is the primary cause of DOCX/HWPX headings appearing as plain paragraphs.

### Building a CDM document: `DocumentBuilder` (`src/cdm/document_builder.hpp`)

All format handlers use `DocumentBuilder` to construct the `cdm::Document`. Call pattern:

```cpp
cdm::DocumentBuilder b;
b.SetOriginalFormat(cdm::FileFormat::DOCX);
b.SetTitle("…"); b.SetAuthor("…");
b.AddStyle(styleDef);        // register named styles
b.BeginSection();
  b.BeginParagraph();
    b.SetCurrentParagraphStyleRef("Heading1");  // sets Paragraph::styleRef
    b.AddText(U"Hello");
  b.EndParagraph();
b.EndSection();
cdm::Document doc = b.MoveBuild();
```

### Normalization (`src/cdm/CdmNormalizer.hpp`)

`cdm::Normalize(Document&)` currently only merges adjacent `Text` inlines with identical `directStyle`. **Missing steps** (high priority):

1. **Resolve style refs** — for each `Paragraph`/`Text`/`Inline` with a `styleRef`, find the matching `StyleDefinition` in `doc.styles` and merge its `TextStyle`/`ParagraphStyle` into `directStyle` (direct style wins on conflict)
2. **Promote heading paragraphs** — paragraphs whose resolved style name matches `"Heading1"`–`"Heading6"` (case-insensitive) should be replaced with `Heading` blocks
3. **Normalize line breaks** — collapse consecutive empty paragraphs, strip trailing whitespace

### Rendering: `CdmLoader` (`src/CdmLoader.h/.cpp`)

Walks the CDM tree, builds a flat `std::wstring` (with `\r\n` = 1 RichEdit position), and collects `CharRun`/`ParaRun` annotations. Then:
1. `editor->SetText(text)` — set plain text
2. `editor->ApplyCharFormat(defaultFont, false)` — SCF_ALL with Malgun Gothic (must come **after** SetText)
3. Apply para formats → apply char formats

Position tracking: each character = 1, each `\r\n` paragraph mark = 1 (RichEdit normalizes `\r\n` → `\r` internally).

### Format handlers (`src/formats/`)

All implement `IDocumentFormat`:

```cpp
FormatResult Load(const std::wstring& path, Document& doc);
FormatResult Save(const std::wstring& path, const std::wstring& content,
                  const std::string& rtfContent, Document& doc);
bool CanWrite() const;
```

`FormatResult::cdmDoc` non-empty sections → CDM path (used by MainWindow). `FormatResult::content` + `rtf=true` → legacy RTF path (still used by some handlers).

| Handler | Read | Write | Notes |
|---------|------|-------|-------|
| TxtFormat | CDM | UTF-8 w/ BOM | Handles UTF-8/16/ANSI detection |
| MdFormat | CDM | plain text | Markdown → CDM blocks; save is plain text |
| HtmlFormat | CDM | full HTML | Skips `<head>`,`<nav>`,`<footer>`,`<aside>` |
| DocxFormat | CDM | ZIP+XML | ZIP via ZipReader; word/document.xml parsing |
| HwpxFormat | CDM | ZIP+XML | ISO/IEC 26300-3 |
| HwpFormat | CDM | ❌ | OLE2 binary; text extraction only |
| DocFormat | CDM | ❌ | OLE2 binary; text extraction only |

### Editor (`src/Editor.h/.cpp`)

Wraps a Win32 RichEdit control (tries `Msftedit.dll` / `RICHEDIT50W` first, falls back to `Riched20.dll`). Key methods used by CdmLoader: `SetText()`, `SetSel(start, end)`, `ApplyCharFormat(CharFormat, selOnly)`, `ApplyParaFormat(ParaFormat, selOnly)`.

`CharFormat::fontSize` is in **half-points** (22 = 11 pt). `ParaFormat::leftIndent` is in **twips** (720 = 0.5 inch).

### MainWindow (`src/MainWindow.cpp`)

Orchestrates file open/save. The CDM loading sequence in `OpenFile()`:

```cpp
cdm::Normalize(result.cdmDoc);                             // merge text runs
LoadCdmDocument(result.cdmDoc, m_editor.get(), textColor); // direct RichEdit load
m_editor->SetBackground(bgColor);
```

`ApplyTheme()` uses `SCF_ALL` which **overwrites all per-character formatting** — never call it after `LoadCdmDocument`.

## Development branch

All work goes to `claude/whats-up-text-editor-OwYMn`. Push with:
```bash
git push -u origin claude/whats-up-text-editor-OwYMn
```
