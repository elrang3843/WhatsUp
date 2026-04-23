#include "HtmlFormat.h"
#include "TxtFormat.h"
#include <algorithm>
#include <sstream>
#include <vector>
#include <cstdint>

static std::wstring ReadTextFile(const std::wstring& path) {
    TxtFormat txt; Document dummy;
    auto r = txt.Load(path, dummy);
    return r.ok ? r.content : L"";
}

// ── RTF helper ────────────────────────────────────────────────────────────────

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

// ── Entity decoder ────────────────────────────────────────────────────────────

static std::wstring DecodeEntities(const std::wstring& text) {
    std::wstring r;
    r.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i) {
        if (text[i] != L'&') { r += text[i]; continue; }
        size_t end = text.find(L';', i + 1);
        if (end == std::wstring::npos || end - i > 10) { r += text[i]; continue; }
        std::wstring ent = text.substr(i + 1, end - i - 1);
        if      (ent == L"amp")  r += L'&';
        else if (ent == L"lt")   r += L'<';
        else if (ent == L"gt")   r += L'>';
        else if (ent == L"quot") r += L'"';
        else if (ent == L"apos") r += L'\'';
        else if (ent == L"nbsp") r += L' ';
        else if (!ent.empty() && ent[0] == L'#') {
            wchar_t ch = static_cast<wchar_t>(
                (ent.size() > 1 && (ent[1] == L'x' || ent[1] == L'X'))
                ? wcstol(ent.c_str() + 2, nullptr, 16)
                : wcstol(ent.c_str() + 1, nullptr, 10));
            r += ch;
        } else { r += L'&'; continue; }
        i = end;
    }
    return r;
}

// ── HTML → RTF converter ──────────────────────────────────────────────────────

static std::string HtmlToRtf(const std::wstring& html) {
    std::string body;

    bool inScript = false, inStyle = false;
    int  boldD = 0, italicD = 0, underD = 0, codeD = 0;
    bool listOl = false;
    int  listNum = 0;

    // Table state
    struct Cell { std::string c; bool isHdr; };
    std::vector<std::vector<Cell>> tblRows;
    std::vector<Cell>              tblRow;
    std::string                    cellBuf;
    bool                           cellIsHdr = false;
    bool                           inCell = false, inRow = false, inTable = false;

    size_t pos = 0;
    while (pos < html.size()) {
        // Text content
        if (html[pos] != L'<') {
            if (!inScript && !inStyle) {
                size_t end = html.find(L'<', pos);
                if (end == std::wstring::npos) end = html.size();
                std::wstring raw = html.substr(pos, end - pos);
                std::wstring decoded = DecodeEntities(raw);
                // Collapse whitespace for non-pre content
                std::wstring norm;
                bool lastWs = false;
                for (wchar_t c : decoded) {
                    if (iswspace(c) && c != L'\n') {
                        if (!lastWs && !norm.empty()) { norm += L' '; lastWs = true; }
                    } else if (c == L'\n') {
                        // ignore raw newlines in HTML (layout is from tags)
                    } else {
                        norm += c; lastWs = false;
                    }
                }
                (inCell ? cellBuf : body) += RtfEnc(norm);
                pos = end;
            } else {
                pos++;
            }
            continue;
        }

        size_t lt = pos;
        size_t gt = html.find(L'>', lt);
        if (gt == std::wstring::npos) { pos++; continue; }

        std::wstring raw = html.substr(lt + 1, gt - lt - 1);
        pos = gt + 1;
        if (raw.empty()) continue;

        // Comment
        if (raw.size() >= 3 && raw.substr(0, 3) == L"!--") {
            size_t ce = html.find(L"-->", lt);
            if (ce != std::wstring::npos) pos = ce + 3;
            continue;
        }

        bool isE  = (raw[0] == L'/');
        bool isSC = (raw.back() == L'/');
        (void)isSC;

        std::wstring nm;
        { size_t i = isE ? 1 : 0;
          while (i < raw.size() && raw[i] != L' ' && raw[i] != L'/' && raw[i] != L'\t')
              nm += towlower(raw[i++]); }

        // Script / style
        if (nm == L"script") { inScript = !isE; continue; }
        if (nm == L"style")  { inStyle  = !isE; continue; }
        if (inScript || inStyle) continue;

        if (!isE) {
            static const int hSz[] = {0,36,28,24,22,20,18};
            if (nm==L"h1"||nm==L"h2"||nm==L"h3"||nm==L"h4"||nm==L"h5"||nm==L"h6") {
                int lv = nm[1] - L'0';
                body += "\\pard\\ql\\b\\fs" + std::to_string(hSz[lv]) + " ";
            }
            else if (nm==L"p"||nm==L"div") { body += "\\pard\\ql "; }
            else if (nm==L"br")             { (inCell ? cellBuf : body) += "\\line "; }
            else if (nm==L"hr")             { body += "\\pard\\ql\\brdrb\\brdrs\\brdrw10 \\par\n\\pard\\ql "; }
            else if (nm==L"li") {
                if (listOl) {
                    listNum++;
                    body += "\\pard\\ql\\li360 " + std::to_string(listNum) + ". ";
                } else {
                    body += "\\pard\\ql\\li360 \\bullet  ";
                }
            }
            else if (nm==L"ul") { listOl = false; }
            else if (nm==L"ol") { listOl = true;  listNum = 0; }
            else if (nm==L"pre"||nm==L"code"||nm==L"tt") {
                (inCell?cellBuf:body) += "{\\f3 "; codeD++;
            }
            else if (nm==L"b"||nm==L"strong") { (inCell?cellBuf:body) += "{\\b ";   boldD++;   }
            else if (nm==L"i"||nm==L"em")     { (inCell?cellBuf:body) += "{\\i ";   italicD++; }
            else if (nm==L"u")                { (inCell?cellBuf:body) += "{\\ul ";  underD++;  }
            else if (nm==L"a")                { (inCell?cellBuf:body) += "{\\ul ";  underD++;  }
            else if (nm==L"table")            { inTable=true; tblRows.clear(); }
            else if (nm==L"tr")               { inRow=true; tblRow.clear(); }
            else if (nm==L"td")               { inCell=true; cellBuf.clear(); cellIsHdr=false; }
            else if (nm==L"th")               { inCell=true; cellBuf.clear(); cellIsHdr=true; }
        } else {
            if (nm==L"h1"||nm==L"h2"||nm==L"h3"||nm==L"h4"||nm==L"h5"||nm==L"h6") {
                body += "\\b0\\fs22\\par\n";
            }
            else if (nm==L"p"||nm==L"div")          { body += "\\par\n"; }
            else if (nm==L"li")                      { body += "\\par\n"; }
            else if (nm==L"pre"||nm==L"code"||nm==L"tt") {
                if (codeD > 0) { (inCell?cellBuf:body) += "}"; codeD--; }
            }
            else if (nm==L"b"||nm==L"strong") {
                if (boldD   > 0) { (inCell?cellBuf:body) += "}"; boldD--;   }
            }
            else if (nm==L"i"||nm==L"em") {
                if (italicD > 0) { (inCell?cellBuf:body) += "}"; italicD--; }
            }
            else if (nm==L"u"||nm==L"a") {
                if (underD  > 0) { (inCell?cellBuf:body) += "}"; underD--;  }
            }
            else if (nm==L"table") {
                if (!tblRows.empty()) {
                    int maxC = 0;
                    for (auto& r : tblRows) maxC = std::max(maxC, (int)r.size());
                    int defW = maxC ? 9360/maxC : 9360;
                    for (auto& row : tblRows) {
                        body += "\\trowd\\trqc\n";
                        int x = 0;
                        for (int c = 0; c < (int)row.size(); c++) {
                            x += defW;
                            body += "\\clbrdrt\\brdrw15\\brdrs"
                                    "\\clbrdrl\\brdrw15\\brdrs"
                                    "\\clbrdrb\\brdrw15\\brdrs"
                                    "\\clbrdrr\\brdrw15\\brdrs";
                            body += "\\cellx" + std::to_string(x) + "\n";
                        }
                        for (auto& cl : row) {
                            body += "\\pard\\intbl ";
                            if (cl.isHdr) body += "{\\b ";
                            body += cl.c;
                            if (cl.isHdr) body += "}";
                            body += "\\cell\n";
                        }
                        body += "\\row\n";
                    }
                    body += "\\pard\\ql\n";
                }
                inTable = false;
            }
            else if (nm==L"tr") {
                if (inRow) { tblRows.push_back(tblRow); inRow=false; }
            }
            else if (nm==L"td"||nm==L"th") {
                if (inCell) { tblRow.push_back({cellBuf, cellIsHdr}); inCell=false; }
            }
        }
    }

    return kRtfHeader + body + "}";
}

// ── WrapHtml (Save) ───────────────────────────────────────────────────────────

std::wstring HtmlFormat::WrapHtml(const std::wstring& text, const DocProperties& props) {
    std::wostringstream ss;
    ss << L"<!DOCTYPE html>\r\n<html lang=\"ko\">\r\n<head>\r\n"
       << L"<meta charset=\"UTF-8\">\r\n"
       << L"<title>" << (props.title.empty() ? L"WhatsUp Document" : props.title) << L"</title>\r\n"
       << L"<style>\r\nbody { font-family: 'Malgun Gothic','Segoe UI',sans-serif; "
          L"font-size:11pt; margin:40px; line-height:1.6; }\r\n</style>\r\n</head>\r\n<body>\r\n";
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        std::wstring esc;
        for (wchar_t ch : line) {
            if      (ch == L'&') esc += L"&amp;";
            else if (ch == L'<') esc += L"&lt;";
            else if (ch == L'>') esc += L"&gt;";
            else                 esc += ch;
        }
        ss << L"<p>" << esc << L"</p>\r\n";
    }
    ss << L"</body>\r\n</html>\r\n";
    return ss.str();
}

// ── Public interface ───────────────────────────────────────────────────────────

FormatResult HtmlFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    std::wstring html = ReadTextFile(path);
    if (html.empty()) { r.error = L"Cannot read HTML file."; return r; }

    std::string rtf = HtmlToRtf(html);
    r.content = std::wstring(rtf.begin(), rtf.end());
    r.rtf     = true;
    r.ok      = true;
    return r;
}

FormatResult HtmlFormat::Save(const std::wstring& path,
                               const std::wstring& content,
                               const std::string&  /*rtf*/,
                               Document&           doc) {
    FormatResult r;
    std::wstring html = WrapHtml(content, doc.Properties());

    int needed = WideCharToMultiByte(CP_UTF8, 0, html.c_str(),
        static_cast<int>(html.size()), nullptr, 0, nullptr, nullptr);
    std::vector<char> utf8(needed);
    WideCharToMultiByte(CP_UTF8, 0, html.c_str(),
        static_cast<int>(html.size()), utf8.data(), needed, nullptr, nullptr);

    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0,
                               nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) { r.error = L"Cannot create file."; return r; }

    const uint8_t bom[3] = { 0xEF, 0xBB, 0xBF };
    DWORD written = 0;
    WriteFile(hFile, bom, 3, &written, nullptr);
    WriteFile(hFile, utf8.data(), needed, &written, nullptr);
    CloseHandle(hFile);
    r.ok = true;
    return r;
}
