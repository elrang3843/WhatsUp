#include "MdFormat.h"
#include "TxtFormat.h"
#include "../cdm/document_builder.hpp"
#include <sstream>
#include <vector>
#include <algorithm>

static std::wstring ReadTextFile(const std::wstring& path) {
    TxtFormat txt; Document dummy;
    auto r = txt.Load(path, dummy);
    // TxtFormat now returns CDM; reconstruct wstring from CDM paragraphs
    if (!r.ok) return {};
    if (!r.cdmDoc.sections.empty()) {
        std::wstring out;
        for (auto& sec : r.cdmDoc.sections) {
            for (auto& blk : sec->blocks) {
                std::visit([&](auto& b) {
                    using T = std::decay_t<decltype(b)>;
                    if constexpr (std::is_same_v<T, cdm::Paragraph>) {
                        for (auto& inl : b.inlines) {
                            if (auto* t = std::get_if<cdm::Text>(&inl->value))
                                for (char32_t c : t->text) out += static_cast<wchar_t>(c);
                        }
                        out += L'\n';
                    }
                }, blk->value);
            }
        }
        if (!out.empty() && out.back() == L'\n') out.pop_back();
        return out;
    }
    return r.content;
}

// ── helpers ───────────────────────────────────────────────────────────────────

static std::u32string ToU32(const std::wstring& ws, size_t from = 0, size_t len = std::wstring::npos) {
    std::u32string s;
    auto end = (len == std::wstring::npos) ? ws.size() : std::min(from + len, ws.size());
    for (size_t i = from; i < end; ++i)
        s += static_cast<char32_t>(static_cast<uint16_t>(ws[i]));
    return s;
}

static bool IsHtmlBlock(const std::wstring& line) {
    return line.size() >= 2 && line[0] == L'<' &&
           (iswalpha(line[1]) || line[1] == L'/' || line[1] == L'!');
}

static bool IsTableSep(const std::wstring& line) {
    if (line.empty() || line[0] != L'|') return false;
    for (wchar_t c : line)
        if (c != L'|' && c != L'-' && c != L':' && c != L' ') return false;
    return line.find(L'-') != std::wstring::npos;
}

static bool IsTableRow(const std::wstring& line) {
    return line.size() >= 2 && line[0] == L'|';
}

static std::vector<std::wstring> SplitTableRow(const std::wstring& line) {
    std::vector<std::wstring> cells;
    size_t i = (line[0] == L'|') ? 1 : 0;
    while (i <= line.size()) {
        size_t next = line.find(L'|', i);
        if (next == std::wstring::npos) next = line.size();
        std::wstring cell = line.substr(i, next - i);
        auto s = cell.find_first_not_of(L" \t");
        auto e = cell.find_last_not_of(L" \t");
        cell = (s != std::wstring::npos) ? cell.substr(s, e - s + 1) : L"";
        cells.push_back(cell);
        i = next + 1;
    }
    if (!cells.empty() && cells.back().empty()) cells.pop_back();
    return cells;
}

// ── Inline MD → CDM ───────────────────────────────────────────────────────────

static void ParseMdInlines(const std::wstring& s, size_t from, size_t to,
                             cdm::DocumentBuilder& b);

static void AddWStr(const std::wstring& ws, size_t from, size_t len,
                    cdm::DocumentBuilder& b)
{
    if (len == 0) return;
    b.AddText(ToU32(ws, from, len));
}

static void ParseMdInlines(const std::wstring& s, size_t from, size_t to,
                             cdm::DocumentBuilder& b)
{
    size_t textStart = from;
    size_t i = from;

    auto flushText = [&]() {
        if (textStart < i) AddWStr(s, textStart, i - textStart, b);
        textStart = i;
    };

    while (i < to) {
        // Bold: **text** or __text__
        if (i + 1 < to &&
            ((s[i]==L'*' && s[i+1]==L'*') || (s[i]==L'_' && s[i+1]==L'_'))) {
            wchar_t d = s[i];
            std::wstring close(2, d);
            size_t end = s.find(close, i + 2);
            if (end != std::wstring::npos && end + 1 < to) {
                flushText();
                b.BeginStrong();
                ParseMdInlines(s, i + 2, end, b);
                b.EndStrong();
                i = end + 2; textStart = i; continue;
            }
        }
        // Italic: *text* or _text_
        if ((s[i]==L'*' || s[i]==L'_') && (i+1 >= to || s[i+1] != s[i])) {
            wchar_t d = s[i];
            size_t end = s.find(d, i + 1);
            if (end != std::wstring::npos && end < to && (end+1 >= to || s[end+1] != d)) {
                flushText();
                b.BeginEmphasis();
                ParseMdInlines(s, i + 1, end, b);
                b.EndEmphasis();
                i = end + 1; textStart = i; continue;
            }
        }
        // Inline code: `text`
        if (s[i] == L'`') {
            size_t end = s.find(L'`', i + 1);
            if (end != std::wstring::npos && end < to) {
                flushText();
                cdm::TextStyle codeStyle;
                codeStyle.fontFamily = "Courier New";
                b.BeginSpan(codeStyle);
                AddWStr(s, i + 1, end - i - 1, b);
                b.EndSpan();
                i = end + 1; textStart = i; continue;
            }
        }
        // Image: ![alt](url) → italic [alt]
        if (s[i] == L'!' && i + 1 < to && s[i+1] == L'[') {
            size_t j = s.find(L']', i + 2);
            if (j != std::wstring::npos && j + 1 < to && s[j+1] == L'(') {
                size_t k = s.find(L')', j + 2);
                if (k != std::wstring::npos) {
                    flushText();
                    if (j > i + 2) {
                        b.BeginEmphasis();
                        b.AddText(U"[");
                        AddWStr(s, i + 2, j - i - 2, b);
                        b.AddText(U"]");
                        b.EndEmphasis();
                    }
                    i = k + 1; textStart = i; continue;
                }
            }
        }
        // Link: [text](url) → underlined text
        if (s[i] == L'[') {
            size_t j = s.find(L']', i + 1);
            if (j != std::wstring::npos && j + 1 < to && s[j+1] == L'(') {
                size_t k = s.find(L')', j + 2);
                if (k != std::wstring::npos) {
                    flushText();
                    b.BeginUnderline();
                    ParseMdInlines(s, i + 1, j, b);
                    b.EndUnderline();
                    i = k + 1; textStart = i; continue;
                }
            }
        }
        // Inline HTML tag: skip
        if (s[i] == L'<' && i + 1 < to &&
            (iswalpha(s[i+1]) || s[i+1] == L'/' || s[i+1] == L'!')) {
            size_t end = s.find(L'>', i + 1);
            if (end != std::wstring::npos && end < to) {
                flushText();
                i = end + 1; textStart = i; continue;
            }
        }
        ++i;
    }
    flushText();
}

// Add a paragraph from an MD line with inline parsing
static void AddMdParagraph(cdm::DocumentBuilder& b, const std::wstring& line) {
    b.BeginParagraph();
    ParseMdInlines(line, 0, line.size(), b);
    b.EndParagraph();
}

// ── Markdown → CDM ───────────────────────────────────────────────────────────

static cdm::Document MdToCdm(const std::wstring& md) {
    cdm::DocumentBuilder b;
    b.SetOriginalFormat(cdm::FileFormat::MD);
    b.BeginSection();

    std::vector<std::wstring> lines;
    {
        std::wistringstream in(md);
        std::wstring l;
        while (std::getline(in, l)) {
            if (!l.empty() && l.back() == L'\r') l.pop_back();
            lines.push_back(std::move(l));
        }
    }

    bool inFence = false;
    std::string fenceCode;
    std::string fenceLang;
    std::vector<std::vector<std::wstring>> tableRows;
    bool inTable = false;

    auto flushTable = [&]() {
        if (tableRows.empty()) { inTable = false; return; }
        b.BeginTable();
        bool isHeader = true;
        for (auto& row : tableRows) {
            b.BeginTableRow();
            for (auto& cell : row)
                b.AddTableCell(ToU32(cell));
            b.EndTableRow();
            (void)isHeader;
            isHeader = false;
        }
        b.EndTable();
        tableRows.clear();
        inTable = false;
    };

    for (const auto& line : lines) {
        // Fenced code block
        if (line.size() >= 3 && line.substr(0, 3) == L"```") {
            if (inTable) flushTable();
            if (!inFence) {
                inFence = true;
                fenceCode.clear();
                std::wstring lang = line.substr(3);
                fenceLang = std::string(lang.begin(), lang.end());
            } else {
                b.AddCodeBlock(fenceCode, fenceLang);
                fenceCode.clear(); fenceLang.clear();
                inFence = false;
            }
            continue;
        }
        if (inFence) {
            for (wchar_t c : line) fenceCode += static_cast<char>(c < 128 ? c : '?');
            fenceCode += '\n';
            continue;
        }

        // Skip standalone HTML block lines
        if (IsHtmlBlock(line)) {
            if (inTable) flushTable();
            continue;
        }

        // Pipe table rows
        if (IsTableRow(line)) {
            if (!IsTableSep(line))
                tableRows.push_back(SplitTableRow(line));
            inTable = true;
            continue;
        } else if (inTable) {
            flushTable();
        }

        // Empty line → empty paragraph
        if (line.empty()) {
            b.AddParagraph(U"");
            continue;
        }

        // ATX headings: # ## ###
        {
            size_t h = 0;
            while (h < line.size() && line[h] == L'#') h++;
            if (h > 0 && h < line.size() && line[h] == L' ') {
                b.AddHeading(static_cast<int>(std::min(h, (size_t)6)),
                             ToU32(line, h + 1));
                continue;
            }
        }

        // Horizontal rule: --- *** ___
        if (line == L"---" || line == L"***" || line == L"___" ||
            line == L"- - -" || line == L"* * *" || line == L"_ _ _") {
            b.AddHorizontalRule();
            continue;
        }

        // Block quote: > text
        if (!line.empty() && line[0] == L'>') {
            b.BeginBlockQuote();
            b.BeginParagraph();
            ParseMdInlines(line, line.size() > 2 ? 2 : line.size(), line.size(), b);
            b.EndParagraph();
            b.EndBlockQuote();
            continue;
        }

        // Unordered list: - * +
        if (line.size() >= 2 &&
            (line[0]==L'-'||line[0]==L'*'||line[0]==L'+') && line[1]==L' ') {
            // Use a single-item list for simplicity (each bullet is its own list)
            b.BeginList(cdm::ListType::Bullet);
            b.AddListItem(ToU32(line, 2));
            b.EndList();
            continue;
        }

        // Ordered list: 1. 2.
        {
            size_t dot = line.find(L'.');
            if (dot != std::wstring::npos && dot > 0 && dot < 4) {
                bool ok = true;
                for (size_t k = 0; k < dot; k++)
                    if (!iswdigit(line[k])) { ok = false; break; }
                if (ok && dot + 1 < line.size() && line[dot+1] == L' ') {
                    b.BeginList(cdm::ListType::Numbered);
                    b.AddListItem(ToU32(line, dot + 2));
                    b.EndList();
                    continue;
                }
            }
        }

        // Normal paragraph with inline formatting
        AddMdParagraph(b, line);
    }

    if (inTable) flushTable();
    if (inFence && !fenceCode.empty()) b.AddCodeBlock(fenceCode, fenceLang);

    b.EndSection();
    return b.MoveBuild();
}

// ── Public interface ───────────────────────────────────────────────────────────

FormatResult MdFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    std::wstring md = ReadTextFile(path);
    if (md.empty()) { r.error = L"Cannot read Markdown file."; return r; }

    r.cdmDoc = MdToCdm(md);
    r.ok     = true;
    return r;
}

FormatResult MdFormat::Save(const std::wstring& path,
                              const std::wstring& content,
                              const std::string&  /*rtf*/,
                              Document& /*doc*/) {
    TxtFormat txt; Document dummy;
    return txt.Save(path, content, "", dummy);
}
