#include "CdmLoader.h"
#include "Editor.h"
#include "cdm/document_model.hpp"
#include <algorithm>
#include <string>
#include <vector>

// Heading sizes in half-points: H1=36pt H2=28pt H3=24pt H4=20pt H5=18pt H6=16pt
static const int kHeadingHalfPt[] = { 72, 56, 48, 40, 36, 32 };

struct CharRun {
    int       start, end;
    CharFormat fmt;
};

struct ParaRun {
    int       start, end;   // [start, end) in RichEdit char positions
    ParaFormat fmt;
};

struct ImagePlacement {
    int             start, end;       // placeholder range in RichEdit positions
    cdm::ResourceId resourceId = 0;
    int             widthPx = 0;
    int             heightPx = 0;
};

// ---- builder ------------------------------------------------------------------

struct CdmLoader {
    Editor*              editor;
    const cdm::Document* doc = nullptr;
    std::wstring         text;
    int                  richPos    = 0;   // mirrors RichEdit internal position (\r\n = 1)
    int                  indentTwips = 0;  // accumulated blockquote indent

    std::vector<CharRun>        charRuns;
    std::vector<ParaRun>        paraRuns;
    std::vector<ImagePlacement> imagePlacements;

    // --- text append helpers --------------------------------------------------

    void AppendChar(wchar_t c) { text += c; ++richPos; }

    void AppendWStr(const std::wstring& s) {
        text += s; richPos += static_cast<int>(s.size());
    }

    void AppendU32(const std::u32string& s) {
        for (char32_t c : s) {
            if (c < 0x10000) {
                AppendChar(static_cast<wchar_t>(c));
            } else {
                // Supplementary plane → surrogate pair (2 wchar_t, 2 RichEdit positions)
                c -= 0x10000;
                AppendChar(static_cast<wchar_t>(0xD800 | (c >> 10)));
                AppendChar(static_cast<wchar_t>(0xDC00 | (c & 0x3FF)));
            }
        }
    }

    // \r\n in WM_SETTEXT is normalized to \r (one char) in RichEdit's internal buffer.
    // So each paragraph mark = 1 RichEdit position.
    void AppendParaMark() { text += L"\r\n"; ++richPos; }

    // --- CDM traversal -------------------------------------------------------

    void ProcessDocument(const cdm::Document& doc) {
        for (auto& sec : doc.sections)
            for (auto& blk : sec->blocks)
                ProcessBlock(*blk);
    }

    void ProcessBlock(const cdm::Block& blk) {
        std::visit([this](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, cdm::Paragraph>)
                ProcessParagraph(v);
            else if constexpr (std::is_same_v<T, cdm::Heading>)
                ProcessHeading(v);
            else if constexpr (std::is_same_v<T, cdm::CodeBlock>)
                ProcessCodeBlock(v);
            else if constexpr (std::is_same_v<T, cdm::HorizontalRule>) {
                int s = richPos;
                AppendWStr(L"────────────────────────────");
                AppendParaMark();
                (void)s;
            }
            else if constexpr (std::is_same_v<T, cdm::Table>)
                ProcessTable(v);
            else if constexpr (std::is_same_v<T, cdm::BlockQuote>)
                ProcessBlockQuote(v);
            else if constexpr (std::is_same_v<T, cdm::ListBlock>)
                ProcessList(v);
            else if constexpr (std::is_same_v<T, cdm::SectionBreakBlock>) {
                int s = richPos;
                AppendWStr(L"════════════════════════════");
                AppendParaMark();
                CharRun cr{}; cr.start = s; cr.end = richPos - 1;
                cr.fmt.textColor = RGB(0xA0, 0xA0, 0xA0);
                charRuns.push_back(cr);
            }
            else if constexpr (std::is_same_v<T, cdm::RawBlock>) {
                int s = richPos;
                AppendWStr(L"[원본 블록]");
                AppendParaMark();
                CharRun cr{}; cr.start = s; cr.end = richPos - 1;
                cr.fmt.italic    = 1;
                cr.fmt.textColor = RGB(0x80, 0x80, 0x80);
                charRuns.push_back(cr);
            }
            else if constexpr (std::is_same_v<T, cdm::CustomBlock>) {
                for (auto& child : v.blocks)
                    if (child) ProcessBlock(*child);
            }
        }, blk.value);
    }

    void ProcessParagraph(const cdm::Paragraph& p) {
        int paraStart = richPos;
        ProcessInlines(p.inlines);
        int paraTextEnd = richPos;
        AppendParaMark();

        // Paragraph-level formatting
        bool needPara = indentTwips > 0 || p.directStyle.alignment;
        if (needPara) {
            ParaRun pr{};
            pr.start = paraStart; pr.end = richPos;
            pr.fmt.leftIndent = indentTwips;
            if (p.directStyle.alignment) {
                switch (*p.directStyle.alignment) {
                    case cdm::Alignment::Center:  pr.fmt.alignment = PFA_CENTER;  break;
                    case cdm::Alignment::Right:   pr.fmt.alignment = PFA_RIGHT;   break;
                    case cdm::Alignment::Justify: pr.fmt.alignment = PFA_JUSTIFY; break;
                    default: break;
                }
            }
            paraRuns.push_back(pr);
        }
        (void)paraTextEnd;
    }

    void ProcessHeading(const cdm::Heading& h) {
        int paraStart = richPos;
        ProcessInlines(h.inlines);
        int paraTextEnd = richPos;
        AppendParaMark();

        // Bold + enlarged font for heading text
        int lvl = std::max(1, std::min(h.level, 6)) - 1;
        CharRun cr{};
        cr.start = paraStart; cr.end = paraTextEnd;
        cr.fmt.bold     = 1;
        cr.fmt.fontSize = kHeadingHalfPt[lvl];
        charRuns.push_back(cr);
    }

    void ProcessCodeBlock(const cdm::CodeBlock& cb) {
        int blockStart = richPos;
        // Split code on newlines
        std::wstring line;
        for (char c : cb.code) {
            if (c == '\n') {
                AppendWStr(line); line.clear();
                AppendParaMark();
            } else {
                line += static_cast<wchar_t>(static_cast<unsigned char>(c));
            }
        }
        if (!line.empty() || !cb.code.empty()) {
            if (!line.empty()) AppendWStr(line);
            AppendParaMark();
        }
        int blockEnd = richPos;

        // Monospace font for entire block
        CharRun cr{};
        cr.start = blockStart; cr.end = blockEnd;
        cr.fmt.fontName = L"Courier New";
        cr.fmt.fontSize = 18;   // 9 pt
        charRuns.push_back(cr);

        // Indent the block
        if (blockStart < blockEnd) {
            ParaRun pr{};
            pr.start = blockStart; pr.end = blockEnd;
            pr.fmt.leftIndent = 360;
            paraRuns.push_back(pr);
        }
    }

    void ProcessBlockQuote(const cdm::BlockQuote& bq) {
        indentTwips += 720;
        for (auto& blk : bq.blocks)
            ProcessBlock(*blk);
        indentTwips -= 720;
    }

    void ProcessList(const cdm::ListBlock& lb) {
        int idx = lb.start;
        for (auto& item : lb.items) {
            int paraStart = richPos;

            // Prefix (embed in text, no PARAFORMAT numbering)
            if (lb.type == cdm::ListType::Bullet) {
                AppendWStr(L"• ");
            } else {
                AppendWStr(std::to_wstring(idx++) + L". ");
            }

            // Item content (take first paragraph only for inline text)
            for (auto& blk : item->blocks) {
                std::visit([this](auto& v) {
                    using T = std::decay_t<decltype(v)>;
                    if constexpr (std::is_same_v<T, cdm::Paragraph>)
                        ProcessInlines(v.inlines);
                }, blk->value);
            }
            AppendParaMark();

            // Apply left indent for each list item paragraph
            ParaRun pr{};
            pr.start = paraStart; pr.end = richPos;
            pr.fmt.leftIndent = 360 + indentTwips;
            paraRuns.push_back(pr);
        }
    }

    void ProcessTable(const cdm::Table& t) {
        for (auto& row : t.rows) {
            bool first = true;
            for (auto& cell : row->cells) {
                if (!first) AppendWStr(L"  |  ");
                first = false;
                for (auto& blk : cell->blocks) {
                    std::visit([this](auto& v) {
                        using T = std::decay_t<decltype(v)>;
                        if constexpr (std::is_same_v<T, cdm::Paragraph>)
                            ProcessInlines(v.inlines);
                    }, blk->value);
                }
            }
            AppendParaMark();
        }
    }

    // --- Inline traversal ----------------------------------------------------

    void ProcessInlines(const std::vector<cdm::InlinePtr>& inlines) {
        for (auto& inl : inlines) ProcessInline(*inl);
    }

    void ProcessInline(const cdm::Inline& inl) {
        std::visit([this](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, cdm::Text>)
                ProcessText(v);
            else if constexpr (std::is_same_v<T, cdm::Tab>)
                AppendChar(L'\t');
            else if constexpr (std::is_same_v<T, cdm::Break>)
                AppendChar(L' ');   // soft line break → space
            else if constexpr (std::is_same_v<T, cdm::InlineCode>) {
                int s = richPos;
                AppendU32(v.text);
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.fontName = L"Courier New"; cr.fmt.fontSize = 18;
                    charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::Strong>) {
                int s = richPos;
                ProcessInlines(v.children);
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.bold = 1; charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::Emphasis>) {
                int s = richPos;
                ProcessInlines(v.children);
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.italic = 1; charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::Underline>) {
                int s = richPos;
                ProcessInlines(v.children);
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.underline = 1; charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::Span>) {
                int s = richPos;
                ProcessInlines(v.children);
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    ApplyStyleToCf(v.directStyle, cr.fmt);
                    charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::Hyperlink>) {
                // URL marker only — display text is in adjacent Text sibling nodes
                // (styled blue+underline by the format parser)
                (void)v;
            }
            else if constexpr (std::is_same_v<T, cdm::Field>) {
                if (!v.resultText.empty())
                    AppendWStr(std::wstring(v.resultText.begin(), v.resultText.end()));
            }
            else if constexpr (std::is_same_v<T, cdm::Image>) {
                int s = richPos;
                if (v.altText.empty())
                    AppendWStr(L"[이미지]");
                else {
                    AppendWStr(L"[이미지: ");
                    // altText is UTF-8; convert to wstring
                    int n = MultiByteToWideChar(CP_UTF8, 0, v.altText.c_str(), -1, nullptr, 0);
                    if (n > 1) {
                        std::wstring ws(n-1, L'\0');
                        MultiByteToWideChar(CP_UTF8, 0, v.altText.c_str(), -1, ws.data(), n);
                        AppendWStr(ws);
                    }
                    AppendWStr(L"]");
                }
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.italic    = 1;
                    cr.fmt.textColor = RGB(0x80, 0x80, 0x80);
                    charRuns.push_back(cr);
                    if (v.resourceId) {
                        ImagePlacement ip{};
                        ip.start      = s;
                        ip.end        = richPos;
                        ip.resourceId = *v.resourceId;
                        // width/height are left at 0 here; B6 will translate
                        // Length → px and populate them.
                        imagePlacements.push_back(ip);
                    }
                }
            }
            else if constexpr (std::is_same_v<T, cdm::NoteRef>) {
                int s = richPos;
                AppendWStr(v.type == cdm::NoteType::Endnote ? L"[미주" : L"[각주");
                if (!v.noteId.empty()) {
                    AppendWStr(L":");
                    std::wstring ws(v.noteId.begin(), v.noteId.end());
                    AppendWStr(ws);
                }
                AppendWStr(L"]");
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.italic    = 1;
                    cr.fmt.textColor = RGB(0x60, 0x60, 0xA0);
                    charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::BookmarkStart>) {
                int s = richPos;
                AppendWStr(L"⟨");  // ⟨
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.textColor = RGB(0xA0, 0xA0, 0xA0);
                    charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::BookmarkEnd>) {
                int s = richPos;
                AppendWStr(L"⟩");  // ⟩
                if (s < richPos) {
                    CharRun cr{}; cr.start = s; cr.end = richPos;
                    cr.fmt.textColor = RGB(0xA0, 0xA0, 0xA0);
                    charRuns.push_back(cr);
                }
            }
            else if constexpr (std::is_same_v<T, cdm::CommentRangeStart> ||
                               std::is_same_v<T, cdm::CommentRangeEnd>) {
                // Comments are handled by a dedicated UI (comment pane), not
                // rendered inline. Explicit branch prevents future silent drops.
                (void)v;
            }
            else if constexpr (std::is_same_v<T, cdm::RawInline>) {
                // Raw fragments may contain arbitrary markup bytes; dumping
                // them into the editor would produce garbage. Strip for now.
                // (data is preserved in the CDM for round-trip save.)
                (void)v;
            }
        }, inl.value);
    }

    void ProcessText(const cdm::Text& t) {
        int s = richPos;
        AppendU32(t.text);
        if (s < richPos && (t.directStyle.bold || t.directStyle.italic ||
                            t.directStyle.underline || t.directStyle.strike ||
                            t.directStyle.color || t.directStyle.backgroundColor ||
                            t.directStyle.fontFamily || t.directStyle.fontSize)) {
            CharRun cr{}; cr.start = s; cr.end = richPos;
            ApplyStyleToCf(t.directStyle, cr.fmt);
            charRuns.push_back(cr);
        }
    }

    void ApplyStyleToCf(const cdm::TextStyle& ds, CharFormat& cf) {
        if (ds.bold)       cf.bold      = *ds.bold ? 1 : 0;
        if (ds.italic)     cf.italic    = *ds.italic ? 1 : 0;
        if (ds.strike)     cf.strikeout = *ds.strike ? 1 : 0;
        if (ds.underline)  cf.underline = (*ds.underline != cdm::UnderlineStyle::None) ? 1 : 0;
        if (ds.color)
            cf.textColor = RGB(ds.color->r, ds.color->g, ds.color->b);
        if (ds.backgroundColor)
            cf.bgColor = RGB(ds.backgroundColor->r, ds.backgroundColor->g, ds.backgroundColor->b);
        if (ds.fontSize) {
            if (ds.fontSize->unit == cdm::Unit::Pt)
                cf.fontSize = static_cast<int>(ds.fontSize->value * 2.0);
            else if (ds.fontSize->unit == cdm::Unit::Px)
                cf.fontSize = static_cast<int>(ds.fontSize->value * 1.5);
        }
        if (ds.fontFamily)
            cf.fontName = std::wstring(ds.fontFamily->begin(), ds.fontFamily->end());
    }

    // --- Apply to RichEdit ---------------------------------------------------

    void Apply(COLORREF textColor) {
        // 1. Set plain text first (WM_SETTEXT clears the control).
        editor->SetText(text);

        // 2. Apply Malgun Gothic + theme colour to ALL text with SCF_ALL.
        //    Must come AFTER SetText so that the new content gets the font.
        {
            CharFormat def{};
            def.fontName  = L"Malgun Gothic";
            def.fontSize  = 22;   // 11 pt = 22 half-points
            def.textColor = textColor;
            editor->ApplyCharFormat(def, false);   // SCF_ALL + SCF_DEFAULT
        }

        // 3. Apply paragraph formats first (before char, so char can override).
        for (auto& pr : paraRuns) {
            editor->SetSel(pr.start, pr.end);
            editor->ApplyParaFormat(pr.fmt, true);
        }

        // 4. Apply character formats.
        for (auto& cr : charRuns) {
            editor->SetSel(cr.start, cr.end);
            editor->ApplyCharFormat(cr.fmt, true);
        }

        // 5. Replace image placeholders with embedded pictures. Iterate in
        //    reverse so replacing an earlier range doesn't shift positions
        //    of later placeholders. On failure (stub / decode error) the
        //    placeholder text is left in place.
        if (doc) {
            for (auto it = imagePlacements.rbegin(); it != imagePlacements.rend(); ++it) {
                const cdm::Resource* res = cdm::FindResource(*doc, it->resourceId);
                if (!res || res->data.empty()) continue;
                editor->InsertImageAt(it->start, it->end, res->data,
                                      it->widthPx, it->heightPx);
            }
        }

        // 6. Deselect and move caret to top.
        editor->SetSel(0, 0);
    }
};

// ---- Public entry point -------------------------------------------------------

void LoadCdmDocument(const cdm::Document& doc, Editor* editor, COLORREF textColor) {
    CdmLoader loader;
    loader.editor = editor;
    loader.doc    = &doc;
    loader.ProcessDocument(doc);
    loader.Apply(textColor);
}
