#include "MdFormat.h"
#include "TxtFormat.h"
#include <sstream>
#include <regex>

static std::wstring ReadTextFile(const std::wstring& path) {
    TxtFormat txt;
    Document  dummy;
    auto r = txt.Load(path, dummy);
    return r.ok ? r.content : L"";
}

std::wstring MdFormat::MdToPlain(const std::wstring& md) {
    // This is a line-by-line Markdown stripping pass (not a full renderer)
    std::wistringstream in(md);
    std::wostringstream out;
    std::wstring line;

    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();

        // ATX headings: # ## ### etc.
        size_t hashes = 0;
        while (hashes < line.size() && line[hashes] == L'#') ++hashes;
        if (hashes > 0 && hashes < line.size() && line[hashes] == L' ') {
            out << line.substr(hashes + 1) << L'\n';
            continue;
        }

        // Horizontal rule: --- *** ___
        if (line == L"---" || line == L"***" || line == L"___" ||
            line == L"- - -" || line == L"* * *") {
            out << L"──────────────────────────────\n";
            continue;
        }

        // Block quote: > text
        if (!line.empty() && line[0] == L'>') {
            out << L"│ " << (line.size() > 2 ? line.substr(2) : L"") << L'\n';
            continue;
        }

        // Unordered list: - * +
        if (line.size() >= 2 && (line[0] == L'-' || line[0] == L'*' || line[0] == L'+')
                && line[1] == L' ') {
            out << L"• " << line.substr(2) << L'\n';
            continue;
        }

        // Ordered list: 1. 2. etc.
        size_t dot = line.find(L'.');
        if (dot != std::wstring::npos && dot > 0 && dot < 4) {
            bool allDigits = true;
            for (size_t k = 0; k < dot; ++k) allDigits &= (iswdigit(line[k]) != 0);
            if (allDigits && dot + 1 < line.size() && line[dot + 1] == L' ') {
                out << line.substr(0, dot + 1) << L" " << line.substr(dot + 2) << L'\n';
                continue;
            }
        }

        // Inline code: strip backtick pairs
        std::wstring stripped;
        for (size_t i = 0; i < line.size(); ) {
            if (line[i] == L'`') {
                // skip matching backticks
                size_t j = line.find(L'`', i + 1);
                if (j != std::wstring::npos) { i = j + 1; continue; }
            }
            // Bold/italic markers: ** * __ _
            if (i + 1 < line.size() && line[i] == L'*' && line[i + 1] == L'*') { i += 2; continue; }
            if (i + 1 < line.size() && line[i] == L'_' && line[i + 1] == L'_') { i += 2; continue; }
            if (line[i] == L'*' || line[i] == L'_') { ++i; continue; }
            // Link: [text](url) → text
            if (line[i] == L'[') {
                size_t j = line.find(L']', i);
                if (j != std::wstring::npos && j + 1 < line.size() && line[j + 1] == L'(') {
                    size_t k = line.find(L')', j + 2);
                    if (k != std::wstring::npos) {
                        stripped += line.substr(i + 1, j - i - 1);
                        i = k + 1;
                        continue;
                    }
                }
            }
            // Image: ![alt](url) → alt
            if (line[i] == L'!' && i + 1 < line.size() && line[i + 1] == L'[') {
                size_t j = line.find(L']', i + 2);
                if (j != std::wstring::npos && j + 1 < line.size() && line[j + 1] == L'(') {
                    size_t k = line.find(L')', j + 2);
                    if (k != std::wstring::npos) {
                        stripped += line.substr(i + 2, j - i - 2);
                        i = k + 1;
                        continue;
                    }
                }
            }
            stripped += line[i++];
        }
        out << stripped << L'\n';
    }
    return out.str();
}

FormatResult MdFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    std::wstring md = ReadTextFile(path);
    if (md.empty()) { r.error = L"Cannot read Markdown file."; return r; }

    r.content = MdToPlain(md);
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult MdFormat::Save(const std::wstring& path,
                             const std::wstring& content,
                             const std::string&  /*rtf*/,
                             Document& /*doc*/) {
    // Save as plain UTF-8 text (caller preserves Markdown markup if editing)
    TxtFormat txt;
    Document  dummy;
    return txt.Save(path, content, "", dummy);
}
