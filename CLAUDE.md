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

First build auto-fetches **zlib** via CMake FetchContent (internet required). All `*.cpp` files under `src/` are picked up by `GLOB_RECURSE` — no CMakeLists.txt edit needed when adding new source files. Output binary: `build/WhatsUp.exe` (MinGW) or `build/Release/WhatsUp.exe` (MSVC).

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

**Critical design rule**: `directStyle` is the inline/direct formatting; `styleRef` is a reference to a named `StyleDefinition` in `doc.styles`. Normalization resolves `styleRef` → merges into `directStyle` (direct wins, named style fills gaps) and promotes paragraphs whose `styleRef` matches a heading name into `Heading` blocks.

**Merge semantics** (in `CdmNormalizer.cpp`):
- *Overlay* — src overwrites dst where src has a value (leaf wins). Used when walking a style's `basedOn` chain root → leaf.
- *Inherit* — src only fills empty slots on dst (dst/direct wins). Used when merging a resolved named style into a node's `directStyle`.

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

`cdm::Normalize(Document&)` is a 4-stage pipeline run over every section. Format handlers are expected to emit raw CDM; normalization is the single place these passes live.

1. **`ResolveBlocks`** — for each `Paragraph`/`Heading`/`Text`/`InlineCode`/`InlineContainer` with a `styleRef`, walk the `basedOn` chain (cycle-guarded), produce a `ResolvedStyle`, and `InheritTextStyle`/`InheritParaStyle` it into `directStyle`. Inline containers also build a *cascade* (`OverlayTextStyle` of resolved + direct) that propagates to children so nested `Text` inherits the container's font/size/color.
2. **`PromoteHeadings`** — `Paragraph` whose `styleRef` (or the referenced definition's `name`) normalizes to `heading1`–`heading6` is rewritten in-place into a `Heading` block with the matching `level`. `title`→1, `subtitle`→2 are also recognized. `NormName` strips spaces/underscores/hyphens and lowercases for matching.
3. **`NormalizeBlocks` / `MergeAdjacentText`** — fold consecutive `Text` inlines that share identical `styleRef` *and* `directStyle` (see `StylesEqual`). Dramatically shrinks DOCX per-character runs.
4. **`CollapseEmptyParagraphs`** — drop runs of consecutive empty `Paragraph`s (empty = no inlines, or only empty `Text`). Any non-`Text` inline (Image, Break, …) means "not empty".

All stages recurse into container blocks: `BlockQuote`, `ListBlock` items, `Table` cells, `CustomBlock`.

### Rendering: `CdmLoader` (`src/CdmLoader.h/.cpp`) — the active renderer

> Note: `src/cdm/CdmRenderer.hpp/.cpp` is an **alternative** RTF-string renderer (`cdm::Document` → RTF for `EM_STREAMIN`). It is not wired into `MainWindow` today — the live path is `CdmLoader` calling the RichEdit API directly. Don't confuse the two when editing styles.



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
| HwpFormat | CDM | ❌ | OLE2 binary; **read-only by policy** (licensing — do not implement Save). Saving is blocked in `FormatManager::Save` and the user is asked to pick another format (HWPX/DOCX/HTML/TXT). |
| DocFormat | CDM | ❌ | OLE2 binary; **read-only by policy** (licensing — do not implement Save). Saving is blocked in `FormatManager::Save` and the user is asked to pick another format (DOCX/HTML/TXT). |

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

### Other subsystems (outside the CDM pipeline)

Not every file is part of the load/render path. These are independent subsystems rooted in `MainWindow`:

- `src/Application.h/.cpp` — app singleton, settings persistence, theme state.
- `src/SplashScreen.h/.cpp` — GDI-drawn startup splash; does not use RichEdit.
- `src/i18n/Localization.h/.cpp` — Korean / English string tables; all user-facing strings go through this.
- `src/spell/SpellChecker.h/.cpp` — wraps Windows 8+ `ISpellCheckerFactory`. Degrades gracefully when unavailable.
- `src/autocomplete/AutoComplete.h/.cpp` — popup word completion fed by document word frequency.
- `src/dialogs/` — modal dialogs (Settings, DocInfo, About, InsertTable, InsertHyperlink, InsertDateTime). New dialogs follow the same `DialogProc` + resource-ID pattern; resource IDs live in `src/resource.h`, resources in `src/WhatsUp.rc`.

When adding a menu item, touch four things together: `WhatsUp.rc` (menu entry + accelerator), `resource.h` (command ID), `MainWindow::OnCommand` (handler), and `Localization` (label text).

## Development branch

Active development branch is `claude/init-project-OVSx8`. Push with:
```bash
git push -u origin claude/init-project-OVSx8
```
(Earlier work landed on `claude/init-project-setup-mB09Z` and `claude/whats-up-text-editor-OwYMn`; run `git branch --show-current` before pushing.)
