#include "HtmlFormat.h"
#include "TxtFormat.h"
#include <algorithm>
#include <sstream>

// Reuse TxtFormat for raw file I/O
static std::wstring ReadTextFile(const std::wstring& path) {
    TxtFormat txt;
    Document  dummy;
    auto r = txt.Load(path, dummy);
    return r.ok ? r.content : L"";
}

std::wstring HtmlFormat::StripTags(const std::wstring& html) {
    std::wstring result;
    result.reserve(html.size());
    bool inTag     = false;
    bool inScript  = false;
    bool inStyle   = false;

    for (size_t i = 0; i < html.size(); ++i) {
        if (html[i] == L'<') {
            // Check for <script or <style
            if (i + 6 < html.size()) {
                std::wstring tag = html.substr(i + 1, 6);
                std::transform(tag.begin(), tag.end(), tag.begin(), ::towlower);
                if (tag.substr(0, 6) == L"script") inScript = true;
                if (tag.substr(0, 5) == L"style") inStyle = true;
                // closing tags
                if (tag.substr(0, 7) == L"/script") inScript = false;
                if (tag.substr(0, 6) == L"/style") inStyle = false;
            }
            inTag = true;
            // Add newline for block elements
            if (i + 1 < html.size()) {
                wchar_t next = html[i + 1];
                std::wstring tagname;
                size_t j = i + 1;
                if (next == L'/') ++j;
                while (j < html.size() && html[j] != L'>' && !iswspace(html[j]))
                    tagname += towlower(html[j++]);
                if (tagname == L"p" || tagname == L"br" || tagname == L"div" ||
                    tagname == L"h1" || tagname == L"h2" || tagname == L"h3" ||
                    tagname == L"h4" || tagname == L"h5" || tagname == L"h6" ||
                    tagname == L"li" || tagname == L"tr")
                    result += L'\n';
            }
            continue;
        }
        if (html[i] == L'>') {
            inTag = false;
            continue;
        }
        if (!inTag && !inScript && !inStyle)
            result += html[i];
    }
    return result;
}

std::wstring HtmlFormat::DecodeEntities(const std::wstring& text) {
    std::wstring r;
    r.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i) {
        if (text[i] != L'&') { r += text[i]; continue; }
        // Find semicolon
        size_t end = text.find(L';', i + 1);
        if (end == std::wstring::npos || end - i > 10) { r += text[i]; continue; }
        std::wstring entity = text.substr(i + 1, end - i - 1);
        if      (entity == L"amp")  r += L'&';
        else if (entity == L"lt")   r += L'<';
        else if (entity == L"gt")   r += L'>';
        else if (entity == L"quot") r += L'"';
        else if (entity == L"apos") r += L'\'';
        else if (entity == L"nbsp") r += L' ';
        else if (entity[0] == L'#') {
            wchar_t ch = static_cast<wchar_t>(
                (entity[1] == L'x' || entity[1] == L'X')
                ? wcstol(entity.c_str() + 2, nullptr, 16)
                : wcstol(entity.c_str() + 1, nullptr, 10));
            r += ch;
        } else { r += L'&'; continue; }
        i = end;
    }
    return r;
}

std::wstring HtmlFormat::WrapHtml(const std::wstring& text, const DocProperties& props) {
    std::wostringstream ss;
    ss << L"<!DOCTYPE html>\r\n<html lang=\"ko\">\r\n<head>\r\n"
       << L"<meta charset=\"UTF-8\">\r\n"
       << L"<meta name=\"author\" content=\"" << props.author << L"\">\r\n"
       << L"<title>" << (props.title.empty() ? L"WhatsUp Document" : props.title) << L"</title>\r\n"
       << L"<style>\r\n"
       << L"body { font-family: 'Malgun Gothic', 'Segoe UI', sans-serif; "
          L"font-size: 11pt; margin: 40px; line-height: 1.6; }\r\n"
       << L"</style>\r\n</head>\r\n<body>\r\n";

    // Convert plain text to paragraphs
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        // Escape HTML special chars
        std::wstring escaped;
        for (wchar_t ch : line) {
            if      (ch == L'&')  escaped += L"&amp;";
            else if (ch == L'<')  escaped += L"&lt;";
            else if (ch == L'>')  escaped += L"&gt;";
            else                  escaped += ch;
        }
        ss << L"<p>" << escaped << L"</p>\r\n";
    }
    ss << L"</body>\r\n</html>\r\n";
    return ss.str();
}

FormatResult HtmlFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    std::wstring html = ReadTextFile(path);
    if (html.empty()) { r.error = L"Cannot read HTML file."; return r; }

    std::wstring text = StripTags(html);
    text = DecodeEntities(text);

    // Collapse multiple blank lines
    std::wstring clean;
    bool prevBlank = false;
    for (size_t i = 0; i < text.size(); ++i) {
        if (text[i] == L'\n' && (i + 1 < text.size() && text[i + 1] == L'\n')) {
            if (!prevBlank) { clean += L'\n'; prevBlank = true; }
        } else {
            clean += text[i];
            prevBlank = false;
        }
    }

    r.content = clean;
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult HtmlFormat::Save(const std::wstring& path,
                               const std::wstring& content,
                               const std::string&  /*rtf*/,
                               Document&           doc) {
    FormatResult r;
    std::wstring html = WrapHtml(content, doc.Properties());

    // Write as UTF-8
    int needed = WideCharToMultiByte(CP_UTF8, 0, html.c_str(),
                                     static_cast<int>(html.size()), nullptr, 0, nullptr, nullptr);
    std::vector<char> utf8(needed);
    WideCharToMultiByte(CP_UTF8, 0, html.c_str(),
                        static_cast<int>(html.size()), utf8.data(), needed, nullptr, nullptr);

    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0,
                               nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) { r.error = L"Cannot create file."; return r; }

    // UTF-8 BOM
    const uint8_t bom[3] = { 0xEF, 0xBB, 0xBF };
    DWORD written = 0;
    WriteFile(hFile, bom, 3, &written, nullptr);
    WriteFile(hFile, utf8.data(), needed, &written, nullptr);
    CloseHandle(hFile);
    r.ok = true;
    return r;
}
