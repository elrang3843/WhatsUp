#include "CdmRenderer.hpp"
#include <algorithm>
#include <sstream>
#include <cstdio>

namespace cdm {

// ---- helpers ----

static uint32_t PackColor(const Color& c) {
    return (static_cast<uint32_t>(c.r) << 16)
         | (static_cast<uint32_t>(c.g) << 8)
         |  static_cast<uint32_t>(c.b);
}

int CdmRenderer::ColorIndex(const Color& c) {
    uint32_t packed = PackColor(c);
    for (int i = 0; i < static_cast<int>(colorTable_.size()); ++i)
        if (colorTable_[i] == packed) return i + 1; // RTF color table is 1-based
    colorTable_.push_back(packed);
    return static_cast<int>(colorTable_.size());
}

int CdmRenderer::FontIndex(const std::string& family) {
    for (int i = 0; i < static_cast<int>(fontTable_.size()); ++i)
        if (fontTable_[i] == family) return i;
    fontTable_.push_back(family);
    return static_cast<int>(fontTable_.size()) - 1;
}

// ---- table builder pass ----

void CdmRenderer::BuildTables(const Document& doc) {
    // Default fonts
    FontIndex("Arial");
    FontIndex("Courier New");
    // Scan all content for colors / fonts (simplified: just ensure defaults exist)
    (void)doc;
}

// ---- RTF escape ----

std::string CdmRenderer::EscapeRtf(const std::u32string& text) {
    std::string out;
    out.reserve(text.size() * 2);
    for (char32_t ch : text) {
        if (ch == U'\\') { out += "\\\\"; }
        else if (ch == U'{')  { out += "\\{";  }
        else if (ch == U'}')  { out += "\\}";  }
        else if (ch < 128)    { out += static_cast<char>(ch); }
        else if (ch < 0x10000) {
            char buf[16];
            std::snprintf(buf, sizeof(buf), "\\u%d?", static_cast<int16_t>(static_cast<int>(ch)));
            out += buf;
        } else {
            // Surrogate pair approximation for RTF
            char buf[32];
            std::snprintf(buf, sizeof(buf), "\\u%d?", static_cast<int>(ch));
            out += buf;
        }
    }
    return out;
}

std::string CdmRenderer::EscapeRtfA(const std::string& text) {
    std::u32string u32;
    u32.reserve(text.size());
    for (unsigned char c : text) u32 += static_cast<char32_t>(c);
    return EscapeRtf(u32);
}

// ---- style application ----

void CdmRenderer::ApplyTextStyle(const TextStyle& s, std::string& out) {
    if (s.bold && *s.bold)   out += "\\b";
    if (s.italic && *s.italic) out += "\\i";
    if (s.underline && *s.underline != UnderlineStyle::None) out += "\\ul";
    if (s.strike && *s.strike) out += "\\strike";
    if (s.color) {
        int idx = const_cast<CdmRenderer*>(this)->ColorIndex(*s.color);
        char buf[16];
        std::snprintf(buf, sizeof(buf), "\\cf%d", idx);
        out += buf;
    }
    if (s.backgroundColor) {
        int idx = const_cast<CdmRenderer*>(this)->ColorIndex(*s.backgroundColor);
        char buf[16];
        std::snprintf(buf, sizeof(buf), "\\highlight%d", idx);
        out += buf;
    }
    if (s.fontSize) {
        int halfPts = 0;
        if (s.fontSize->unit == Unit::Pt)
            halfPts = static_cast<int>(s.fontSize->value * 2.0);
        else if (s.fontSize->unit == Unit::Px)
            halfPts = static_cast<int>(s.fontSize->value * 1.5);
        if (halfPts > 0) {
            char buf[16];
            std::snprintf(buf, sizeof(buf), "\\fs%d", halfPts);
            out += buf;
        }
    }
    if (s.fontFamily) {
        int idx = const_cast<CdmRenderer*>(this)->FontIndex(*s.fontFamily);
        char buf[16];
        std::snprintf(buf, sizeof(buf), "\\f%d", idx);
        out += buf;
    }
    if (!out.empty()) out += ' ';
}

void CdmRenderer::ResetStyle(std::string& out) {
    out += "\\plain ";
}

// ---- main render ----

std::string CdmRenderer::Render(const Document& doc) {
    colorTable_.clear();
    fontTable_.clear();
    indentLeft_ = 0;

    // Pre-populate default fonts so they get stable indices
    FontIndex("Malgun Gothic");  // f0 – default Korean/body font
    FontIndex("Courier New");    // f1 – monospace

    // Pass 1: render body (this discovers all colors and additional fonts)
    rtf_.clear();
    WriteDocument(doc);
    std::string body = std::move(rtf_);

    // Pass 2: write final RTF = header + body + close
    rtf_.clear();
    WriteHeader();
    rtf_ += body;
    rtf_ += '}';
    return rtf_;
}

void CdmRenderer::WriteHeader() {
    // \ansicpg1252: safe Western codepage; all non-ASCII uses \u escapes anyway
    // \uc1: each \uN escape has 1 fallback character (the '?')
    rtf_ += "{\\rtf1\\ansi\\ansicpg1252\\uc1\\deff0";

    // Font table — charset 0 (ANSI) for all; Unicode escapes handle non-Latin chars
    rtf_ += "{\\fonttbl";
    for (int i = 0; i < static_cast<int>(fontTable_.size()); ++i) {
        char buf[256];
        std::snprintf(buf, sizeof(buf),
            "{\\f%d\\fswiss\\fcharset0 %s;}", i, fontTable_[i].c_str());
        rtf_ += buf;
    }
    rtf_ += '}';

    // Color table (1-based; entry at index 0 is the implicit "auto" color)
    if (!colorTable_.empty()) {
        rtf_ += "{\\colortbl;";
        for (uint32_t c : colorTable_) {
            char buf[64];
            std::snprintf(buf, sizeof(buf), "\\red%u\\green%u\\blue%u;",
                (c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF);
            rtf_ += buf;
        }
        rtf_ += '}';
    }

    // Default paragraph + font settings
    rtf_ += "\\f0\\fs22\\pard ";
}

void CdmRenderer::WriteDocument(const Document& doc) {
    for (auto& sec : doc.sections)
        WriteSection(*sec);
}

void CdmRenderer::WriteSection(const Section& sec) {
    for (auto& blk : sec.blocks)
        WriteBlock(*blk);
}

void CdmRenderer::WriteBlock(const Block& blk) {
    std::visit([this](auto& b) {
        using T = std::decay_t<decltype(b)>;
        if constexpr (std::is_same_v<T, Paragraph>)     WriteParagraph(b);
        else if constexpr (std::is_same_v<T, Heading>)  WriteHeading(b);
        else if constexpr (std::is_same_v<T, CodeBlock>) WriteCodeBlock(b);
        else if constexpr (std::is_same_v<T, HorizontalRule>) WriteHRule();
        else if constexpr (std::is_same_v<T, Table>)    WriteTable(b);
        else if constexpr (std::is_same_v<T, BlockQuote>) WriteBlockQuote(b);
        else if constexpr (std::is_same_v<T, ListBlock>) WriteList(b);
        else if constexpr (std::is_same_v<T, RawBlock>) {
            // skip raw blocks (HTML etc.)
        }
        else if constexpr (std::is_same_v<T, SectionBreakBlock>) {
            rtf_ += "\\page\n";
        }
    }, blk.value);
}

void CdmRenderer::WriteParagraph(const Paragraph& p, bool isList) {
    rtf_ += "\\pard";
    int li = (isList ? 720 : 0) + indentLeft_;
    if (li > 0) {
        char buf[16];
        std::snprintf(buf, sizeof(buf), "\\li%d", li);
        rtf_ += buf;
    }
    if (p.directStyle.alignment) {
        switch (*p.directStyle.alignment) {
            case Alignment::Center:  rtf_ += "\\qc"; break;
            case Alignment::Right:   rtf_ += "\\qr"; break;
            case Alignment::Justify: rtf_ += "\\qj"; break;
            default: break;
        }
    }
    rtf_ += ' ';
    WriteInlines(p.inlines);
    rtf_ += "\\par\n";
}

void CdmRenderer::WriteHeading(const Heading& h) {
    rtf_ += "\\pard";
    // Heading sizes: H1=36pt, H2=28pt, H3=24pt, H4=20pt, H5=18pt, H6=16pt (half-points)
    static const int sizes[] = {72, 56, 48, 40, 36, 32};
    int lvl = std::max(1, std::min(h.level, 6)) - 1;
    char buf[32];
    std::snprintf(buf, sizeof(buf), "\\fs%d\\b ", sizes[lvl]);
    rtf_ += buf;
    WriteInlines(h.inlines);
    rtf_ += "\\b0\\fs22\\par\n";
}

void CdmRenderer::WriteCodeBlock(const CodeBlock& cb) {
    rtf_ += "\\pard\\f1\\fs20 ";
    rtf_ += EscapeRtfA(cb.code);
    rtf_ += "\\f0\\fs22\\par\n";
}

void CdmRenderer::WriteHRule() {
    // Horizontal rule: thin border below empty paragraph
    rtf_ += "\\pard\\brdrb\\brdrs\\brdrw10\\brsp20 \\par\n";
}

void CdmRenderer::WriteTable(const Table& t) {
    if (t.rows.empty()) return;

    // Determine column count from first row
    size_t cols = 0;
    for (auto& row : t.rows)
        cols = std::max(cols, row->cells.size());
    if (cols == 0) return;

    // Fixed column width: 1440 twips = 1 inch per column
    const int colWidth = 1440;

    for (size_t ri = 0; ri < t.rows.size(); ++ri) {
        auto& row = *t.rows[ri];
        bool isHeader = (ri == 0);

        rtf_ += "\\trowd\\trgaph108\\trleft0";
        // Cell definitions
        for (size_t ci = 0; ci < cols; ++ci) {
            char buf[64];
            std::snprintf(buf, sizeof(buf),
                "\\clbrdrt\\brdrs\\brdrw10"
                "\\clbrdrl\\brdrs\\brdrw10"
                "\\clbrdrb\\brdrs\\brdrw10"
                "\\clbrdrr\\brdrs\\brdrw10"
                "\\cellx%d",
                static_cast<int>((ci + 1) * colWidth));
            rtf_ += buf;
        }

        // Cell contents
        for (size_t ci = 0; ci < cols; ++ci) {
            rtf_ += "\\pard\\intbl";
            if (isHeader) rtf_ += "\\b";
            rtf_ += ' ';
            if (ci < row.cells.size()) {
                for (auto& blk : row.cells[ci]->blocks)
                    std::visit([this](auto& b) {
                        using T = std::decay_t<decltype(b)>;
                        if constexpr (std::is_same_v<T, Paragraph>)
                            WriteInlines(b.inlines);
                    }, blk->value);
            }
            rtf_ += "\\cell\n";
        }
        rtf_ += "\\row\n";
    }
    // Return to normal paragraph mode
    rtf_ += "\\pard\\par\n";
}

void CdmRenderer::WriteBlockQuote(const BlockQuote& bq) {
    indentLeft_ += 720;
    for (auto& blk : bq.blocks)
        WriteBlock(*blk);
    indentLeft_ -= 720;
}

void CdmRenderer::WriteList(const ListBlock& lb) {
    int idx = lb.start;
    for (auto& item : lb.items) {
        rtf_ += "\\pard\\li720 ";
        // Bullet / number prefix
        if (lb.type == ListType::Bullet) {
            rtf_ += "\\bullet  ";
        } else {
            char buf[16];
            std::snprintf(buf, sizeof(buf), "%d. ", idx++);
            rtf_ += buf;
        }
        for (auto& blk : item->blocks)
            std::visit([this](auto& b) {
                using T = std::decay_t<decltype(b)>;
                if constexpr (std::is_same_v<T, Paragraph>)
                    WriteInlines(b.inlines);
            }, blk->value);
        rtf_ += "\\par\n";
    }
}

void CdmRenderer::WriteInlines(const std::vector<InlinePtr>& inlines) {
    for (auto& inl : inlines)
        WriteInline(*inl);
}

void CdmRenderer::WriteInline(const Inline& inl) {
    std::visit([this](auto& v) {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, Text>)       WriteText(v);
        else if constexpr (std::is_same_v<T, Tab>)   rtf_ += "\\tab ";
        else if constexpr (std::is_same_v<T, Break>) {
            if (v.type == BreakType::LineBreak || v.type == BreakType::SoftBreak)
                rtf_ += "\\line ";
            else
                rtf_ += "\\page ";
        }
        else if constexpr (std::is_same_v<T, InlineCode>) {
            rtf_ += "\\f1\\fs18 ";
            rtf_ += EscapeRtf(v.text);
            rtf_ += "\\f0\\fs22 ";
        }
        else if constexpr (std::is_same_v<T, Strong>)
            WriteInlineContainer(v, "\\b ",    "\\b0 ");
        else if constexpr (std::is_same_v<T, Emphasis>)
            WriteInlineContainer(v, "\\i ",    "\\i0 ");
        else if constexpr (std::is_same_v<T, Underline>)
            WriteInlineContainer(v, "\\ul ",   "\\ulnone ");
        else if constexpr (std::is_same_v<T, Span>) {
            // Span needs full style — use group only when font changes (unavoidable)
            bool needsFontChange = v.directStyle.fontFamily || v.directStyle.fontSize;
            if (needsFontChange) {
                std::string grp = "{";
                ApplyTextStyle(v.directStyle, grp);
                rtf_ += grp;
                WriteInlines(v.children);
                rtf_ += '}';
            } else {
                // toggle style, no group
                std::string on, off;
                if (v.directStyle.bold && *v.directStyle.bold)
                    { on += "\\b "; off += "\\b0 "; }
                if (v.directStyle.italic && *v.directStyle.italic)
                    { on += "\\i "; off += "\\i0 "; }
                if (v.directStyle.underline &&
                    *v.directStyle.underline != UnderlineStyle::None)
                    { on += "\\ul "; off += "\\ulnone "; }
                if (v.directStyle.color) {
                    char buf[16];
                    std::snprintf(buf, sizeof(buf), "\\cf%d ",
                                  ColorIndex(*v.directStyle.color));
                    on += buf; off += "\\cf0 ";
                }
                rtf_ += on;
                WriteInlines(v.children);
                rtf_ += off;
            }
        }
        else if constexpr (std::is_same_v<T, Hyperlink>) {
            // Display text is a sibling Text node; Hyperlink just carries href metadata
            (void)v;
        }
        else if constexpr (std::is_same_v<T, Image>) {
            rtf_ += "[image]";
        }
        // other inline types: skip
    }, inl.value);
}

void CdmRenderer::WriteText(const Text& t) {
    const auto& s = t.directStyle;
    bool hasStyle = s.bold || s.italic || s.underline || s.color ||
                    s.fontFamily || s.fontSize || s.strike;
    if (hasStyle) {
        // Emit toggle-style formatting (no groups) to avoid RTF parser issues
        if (s.bold && *s.bold)         rtf_ += "\\b";
        if (s.italic && *s.italic)     rtf_ += "\\i";
        if (s.underline && *s.underline != UnderlineStyle::None) rtf_ += "\\ul";
        if (s.strike && *s.strike)     rtf_ += "\\strike";
        if (s.color) {
            char buf[16];
            std::snprintf(buf, sizeof(buf), "\\cf%d", ColorIndex(*s.color));
            rtf_ += buf;
        }
        if (s.fontSize) {
            int hp = 0;
            if (s.fontSize->unit == Unit::Pt)
                hp = static_cast<int>(s.fontSize->value * 2.0);
            else if (s.fontSize->unit == Unit::Px)
                hp = static_cast<int>(s.fontSize->value * 1.5);
            if (hp > 0) {
                char buf[16];
                std::snprintf(buf, sizeof(buf), "\\fs%d", hp);
                rtf_ += buf;
            }
        }
        if (s.fontFamily) {
            char buf[16];
            std::snprintf(buf, sizeof(buf), "\\f%d", FontIndex(*s.fontFamily));
            rtf_ += buf;
        }
        rtf_ += ' ';
        rtf_ += EscapeRtf(t.text);
        // Reset to paragraph defaults
        if (s.bold && *s.bold)         rtf_ += "\\b0";
        if (s.italic && *s.italic)     rtf_ += "\\i0";
        if (s.underline && *s.underline != UnderlineStyle::None) rtf_ += "\\ulnone";
        if (s.strike && *s.strike)     rtf_ += "\\strike0";
        if (s.color)                   rtf_ += "\\cf0";
        if (s.fontSize)                rtf_ += "\\fs22";
        if (s.fontFamily)              rtf_ += "\\f0";
        rtf_ += ' ';
    } else {
        rtf_ += EscapeRtf(t.text);
    }
}

void CdmRenderer::WriteInlineContainer(const InlineContainer& ic,
    const char* onCmd, const char* offCmd)
{
    rtf_ += onCmd;
    WriteInlines(ic.children);
    rtf_ += offCmd;
}

} // namespace cdm
