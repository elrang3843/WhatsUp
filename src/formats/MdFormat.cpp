#include "MdFormat.h"
#include "TxtFormat.h"
#include <sstream>
#include <vector>
#include <algorithm>

static std::wstring ReadTextFile(const std::wstring& path) {
    TxtFormat txt; Document dummy;
    auto r = txt.Load(path, dummy);
    return r.ok ? r.content : L"";
}

// ── RTF helpers ───────────────────────────────────────────────────────────────

static std::string RtfEnc(const std::wstring& ws) {
    std::string s;
    s.reserve(ws.size() * 2);
    for (wchar_t c : ws) {
        if      (c == L'\\') s += "\\\\";
        else if (c == L'{')  s += "\\{";
        else if (c == L'}')  s += "\\}";
        else if (c == L'\r') {}
        else if (c < 128)    s += static_cast<char>(c);
        else {
            s += "\\u";
            s += std::to_string(static_cast<int>(static_cast<int16_t>(c)));
            s += "?";
        }
    }
    return s;
}

static const std::string kRtfHeader =
    "{\\rtf1\\ansi\\deff2\n"
    "{\\fonttbl\n"
    "{\\f0\\froman\\fcharset0 Times New Roman;}\n"
    "{\\f1\\fswiss\\fcharset0 Calibri;}\n"
    "{\\f2\\fswiss\\fcharset129 Malgun Gothic;}\n"
    "{\\f3\\fmodern\\fcharset0 Courier New;}\n"
    "}\n"
    "\\f2\\fs22\\pard\\ql\n";

// ── HTML helpers ───────────────────────────────────────────────────────────────

// A line is an HTML block if it starts with '<' followed by a letter, '/', or '!'
static bool IsHtmlBlock(const std::wstring& line) {
    return line.size() >= 2 && line[0] == L'<' &&
           (iswalpha(line[1]) || line[1] == L'/' || line[1] == L'!');
}

// ── Table helpers ──────────────────────────────────────────────────────────────

// Separator row: only |, -, :, and spaces
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

// ── Inline Markdown → RTF ─────────────────────────────────────────────────────

std::string MdFormat::InlineToRtf(const std::wstring& line) {
    std::string out;
    for (size_t i = 0; i < line.size(); ) {
        // Bold: **text** or __text__
        if (i + 1 < line.size() &&
            ((line[i]==L'*' && line[i+1]==L'*') || (line[i]==L'_' && line[i+1]==L'_'))) {
            wchar_t d = line[i];
            std::wstring close(2, d);
            size_t end = line.find(close, i + 2);
            if (end != std::wstring::npos) {
                out += "{\\b " + InlineToRtf(line.substr(i+2, end-i-2)) + "}";
                i = end + 2; continue;
            }
        }
        // Italic: *text* or _text_ (single)
        if ((line[i]==L'*' || line[i]==L'_') &&
            (i+1 >= line.size() || line[i+1] != line[i])) {
            wchar_t d = line[i];
            size_t end = line.find(d, i + 1);
            if (end != std::wstring::npos && (end+1 >= line.size() || line[end+1] != d)) {
                out += "{\\i " + InlineToRtf(line.substr(i+1, end-i-1)) + "}";
                i = end + 1; continue;
            }
        }
        // Inline code: `text`
        if (line[i] == L'`') {
            size_t end = line.find(L'`', i + 1);
            if (end != std::wstring::npos) {
                out += "{\\f3 " + RtfEnc(line.substr(i+1, end-i-1)) + "}";
                i = end + 1; continue;
            }
        }
        // Image: ![alt](url) → italic [alt] (can't embed remote images)
        if (line[i] == L'!' && i+1 < line.size() && line[i+1] == L'[') {
            size_t j = line.find(L']', i + 2);
            if (j != std::wstring::npos && j+1 < line.size() && line[j+1] == L'(') {
                size_t k = line.find(L')', j + 2);
                if (k != std::wstring::npos) {
                    std::wstring alt = line.substr(i+2, j-i-2);
                    if (!alt.empty())
                        out += "{\\i [" + RtfEnc(alt) + "]}";
                    i = k + 1; continue;
                }
            }
        }
        // Link: [text](url) → underlined text
        if (line[i] == L'[') {
            size_t j = line.find(L']', i + 1);
            if (j != std::wstring::npos && j+1 < line.size() && line[j+1] == L'(') {
                size_t k = line.find(L')', j + 2);
                if (k != std::wstring::npos) {
                    out += "{\\ul " + InlineToRtf(line.substr(i+1, j-i-1)) + "}";
                    i = k + 1; continue;
                }
            }
        }
        // Inline HTML tag: skip (letter/slash/bang after '<')
        if (line[i] == L'<' && i+1 < line.size() &&
            (iswalpha(line[i+1]) || line[i+1] == L'/' || line[i+1] == L'!')) {
            size_t end = line.find(L'>', i + 1);
            if (end != std::wstring::npos) {
                i = end + 1; continue;
            }
        }
        // Regular character
        wchar_t ch = line[i++];
        if      (ch == L'\\') out += "\\\\";
        else if (ch == L'{')  out += "\\{";
        else if (ch == L'}')  out += "\\}";
        else if (ch < 128)    out += static_cast<char>(ch);
        else {
            out += "\\u";
            out += std::to_string(static_cast<int>(static_cast<int16_t>(ch)));
            out += "?";
        }
    }
    return out;
}

// ── Markdown → RTF ────────────────────────────────────────────────────────────

std::string MdFormat::MdToRtf(const std::wstring& md) {
    std::string body;

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
    std::vector<std::vector<std::wstring>> tableRows;
    bool inTable = false;

    auto flushTable = [&]() {
        if (tableRows.empty()) { inTable = false; return; }
        size_t cols = 0;
        for (auto& r : tableRows) cols = std::max(cols, r.size());
        if (cols > 0) {
            int cellW = 9000 / static_cast<int>(cols);
            bool hdr = true;
            for (auto& row : tableRows) {
                body += "\\trowd\\trgaph108\\trleft0\n";
                int cumX = 0;
                for (size_t c = 0; c < cols; c++) {
                    cumX += cellW;
                    body += "\\clbrdrt\\brdrs\\brdrw5"
                            "\\clbrdrb\\brdrs\\brdrw5"
                            "\\clbrdrl\\brdrs\\brdrw5"
                            "\\clbrdrr\\brdrs\\brdrw5\\cellx";
                    body += std::to_string(cumX);
                }
                body += "\n";
                for (size_t c = 0; c < cols; c++) {
                    const std::wstring& text = (c < row.size()) ? row[c] : L"";
                    body += "\\pard\\intbl";
                    if (hdr) body += "\\b";
                    body += " ";
                    body += InlineToRtf(text);
                    if (hdr) body += "\\b0";
                    body += "\\cell\n";
                }
                body += "\\row\n";
                hdr = false;
            }
            body += "\\pard\\ql\n";
        }
        tableRows.clear();
        inTable = false;
    };

    for (const auto& line : lines) {
        // Fenced code block
        if (line.size() >= 3 && line.substr(0, 3) == L"```") {
            if (inTable) flushTable();
            if (!inFence) { body += "\\pard\\ql\\f3 "; inFence = true; }
            else          { body += "\\f2\\fs22\\par\n"; inFence = false; }
            continue;
        }
        if (inFence) { body += RtfEnc(line) + "\\line\n"; continue; }

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

        // Empty line → paragraph break
        if (line.empty()) { body += "\\pard\\ql\\par\n"; continue; }

        // ATX headings: # ## ###
        size_t h = 0;
        while (h < line.size() && line[h] == L'#') h++;
        if (h > 0 && h < line.size() && line[h] == L' ') {
            static const int hSz[] = {0,36,28,24,22,20,18};
            int sz = hSz[std::min(h, (size_t)6)];
            body += "\\pard\\ql\\b\\fs" + std::to_string(sz) + " ";
            body += InlineToRtf(line.substr(h + 1));
            body += "\\b0\\fs22\\par\n";
            continue;
        }

        // Horizontal rule: --- *** ___
        if (line == L"---" || line == L"***" || line == L"___" ||
            line == L"- - -" || line == L"* * *" || line == L"_ _ _") {
            body += "\\pard\\ql\\brdrb\\brdrs\\brdrw10 \\par\n";
            continue;
        }

        // Block quote: > text
        if (!line.empty() && line[0] == L'>') {
            body += "\\pard\\ql\\li360 ";
            body += InlineToRtf(line.size() > 2 ? line.substr(2) : L"");
            body += "\\par\n";
            continue;
        }

        // Unordered list: - * +
        if (line.size() >= 2 &&
            (line[0]==L'-'||line[0]==L'*'||line[0]==L'+') && line[1]==L' ') {
            body += "\\pard\\ql\\fi-180\\li360 \\bullet  ";
            body += InlineToRtf(line.substr(2));
            body += "\\par\n";
            continue;
        }

        // Ordered list: 1. 2.
        {
            size_t dot = line.find(L'.');
            if (dot != std::wstring::npos && dot > 0 && dot < 4) {
                bool ok = true;
                for (size_t k = 0; k < dot; k++)
                    if (!iswdigit(line[k])) { ok = false; break; }
                if (ok && dot+1 < line.size() && line[dot+1] == L' ') {
                    body += "\\pard\\ql\\li360 ";
                    body += RtfEnc(line.substr(0, dot+1)) + " ";
                    body += InlineToRtf(line.substr(dot + 2));
                    body += "\\par\n";
                    continue;
                }
            }
        }

        // Normal paragraph
        body += "\\pard\\ql ";
        body += InlineToRtf(line);
        body += "\\par\n";
    }

    if (inTable) flushTable();

    return kRtfHeader + body + "}";
}

// ── Public interface ───────────────────────────────────────────────────────────

FormatResult MdFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    std::wstring md = ReadTextFile(path);
    if (md.empty()) { r.error = L"Cannot read Markdown file."; return r; }

    std::string rtf = MdToRtf(md);
    r.content = std::wstring(rtf.begin(), rtf.end());
    r.rtf     = true;
    r.ok      = true;
    return r;
}

FormatResult MdFormat::Save(const std::wstring& path,
                             const std::wstring& content,
                             const std::string&  /*rtf*/,
                             Document& /*doc*/) {
    TxtFormat txt; Document dummy;
    return txt.Save(path, content, "", dummy);
}
