#include "HtmlFormat.h"
#include "../cdm/document_builder.hpp"
#include <algorithm>
#include <sstream>
#include <vector>
#include <cstdint>
#include <map>

static std::wstring ReadFileAsText(const std::wstring& path) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
                           nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return {};
    DWORD sz = GetFileSize(h, nullptr);
    std::vector<uint8_t> buf(sz);
    DWORD rd = 0;
    ReadFile(h, buf.data(), sz, &rd, nullptr);
    CloseHandle(h);
    buf.resize(rd);
    if (buf.size() >= 2 && buf[0] == 0xFF && buf[1] == 0xFE)
        return std::wstring(reinterpret_cast<const wchar_t*>(buf.data() + 2),
                            (buf.size() - 2) / 2);
    size_t start = (buf.size() >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF) ? 3 : 0;
    int needed = MultiByteToWideChar(CP_UTF8, 0,
        reinterpret_cast<LPCCH>(buf.data() + start),
        static_cast<int>(buf.size() - start), nullptr, 0);
    if (needed <= 0) return {};
    std::wstring ws(needed, 0);
    MultiByteToWideChar(CP_UTF8, 0,
        reinterpret_cast<LPCCH>(buf.data() + start),
        static_cast<int>(buf.size() - start), ws.data(), needed);
    return ws;
}

// ── Color helpers ─────────────────────────────────────────────────────────────

// Returns 6-char uppercase hex RGB string from a CSS color value, or empty
static std::wstring ParseCssColor(std::wstring v) {
    size_t s = v.find_first_not_of(L" \t\r\n");
    size_t e = v.find_last_not_of(L" \t\r\n");
    if (s == std::wstring::npos) return {};
    v = v.substr(s, e - s + 1);
    if (!v.empty() && v[0] == L'#') {
        if (v.size() == 7) {
            std::wstring hex = v.substr(1);
            for (auto& c : hex) c = towupper(c);
            return hex;
        }
        if (v.size() == 4) {
            std::wstring hex;
            hex += v[1]; hex += v[1]; hex += v[2]; hex += v[2]; hex += v[3]; hex += v[3];
            for (auto& c : hex) c = towupper(c);
            return hex;
        }
    }
    std::wstring vl = v; for (auto& c : vl) c = towlower(c);
    static const struct { const wchar_t* n; const wchar_t* h; } tbl[] = {
        {L"red",L"FF0000"},{L"green",L"008000"},{L"blue",L"0000FF"},
        {L"yellow",L"FFFF00"},{L"orange",L"FFA500"},{L"purple",L"800080"},
        {L"navy",L"000080"},{L"teal",L"008080"},{L"maroon",L"800000"},
        {L"olive",L"808000"},{L"gray",L"808080"},{L"grey",L"808080"},
        {L"cyan",L"00FFFF"},{L"aqua",L"00FFFF"},{L"magenta",L"FF00FF"},
        {L"fuchsia",L"FF00FF"},{L"lime",L"00FF00"},{L"silver",L"C0C0C0"},
        {L"black",L"000000"},{L"white",L"FFFFFF"},{L"brown",L"A52A2A"},
        {L"pink",L"FFC0CB"},{L"coral",L"FF7F50"},{L"gold",L"FFD700"},
    };
    for (auto& e2 : tbl) if (vl == e2.n) return e2.h;
    return {};
}

// Extract value of a CSS property from a style attribute string
static std::wstring CssProp(const std::wstring& style, const std::wstring& prop) {
    std::wstring sl = style; for (auto& c : sl) c = towlower(c);
    size_t pos = sl.find(prop + L":");
    if (pos == std::wstring::npos) return {};
    pos += prop.size() + 1;
    while (pos < sl.size() && sl[pos] == L' ') pos++;
    size_t end = sl.find(L';', pos);
    if (end == std::wstring::npos) end = sl.size();
    return style.substr(pos, end - pos);
}

// Get value of an HTML attribute (case-insensitive attribute name)
static std::wstring HtmlAttr(const std::wstring& tag, const std::wstring& attr) {
    std::wstring tl = tag; for (auto& c : tl) c = towlower(c);
    std::wstring al = attr; for (auto& c : al) c = towlower(c);
    size_t p = tl.find(al + L"=\"");
    if (p == std::wstring::npos) p = tl.find(al + L"='");
    if (p == std::wstring::npos) return {};
    wchar_t q = tag[p + attr.size() + 1];
    p += attr.size() + 2;
    size_t e = tag.find(q, p);
    return e != std::wstring::npos ? tag.substr(p, e - p) : std::wstring{};
}

static std::string BuildHtmlRtfHeader(const std::map<std::wstring,int>& colorMap) {
    std::string h =
        "{\\rtf1\\ansi\\deff2\n"
        "{\\fonttbl\n"
        "{\\f0\\froman\\fcharset0 Times New Roman;}\n"
        "{\\f1\\fswiss\\fcharset0 Calibri;}\n"
        "{\\f2\\fswiss\\fcharset129 Malgun Gothic;}\n"
        "{\\f3\\fmodern\\fcharset0 Courier New;}\n"
        "}\n"
        "{\\colortbl;\\red0\\green0\\blue0;"; // index 1 = theme placeholder

    // Emit entries in index order (not map's alphabetical order)
    std::vector<std::pair<int,const std::wstring*>> ordered;
    for (auto& [hex, idx] : colorMap)
        ordered.push_back({idx, &hex});
    std::sort(ordered.begin(), ordered.end(),
              [](const auto& a, const auto& b){ return a.first < b.first; });

    for (auto& [sortIdx, phex] : ordered) {
        (void)sortIdx;
        const std::wstring& hex = *phex;
        unsigned r = 0, g = 0, b = 0;
        if (hex.size() == 6) {
            r = static_cast<unsigned>(std::stoul(std::string(hex.begin(), hex.begin()+2), nullptr, 16));
            g = static_cast<unsigned>(std::stoul(std::string(hex.begin()+2, hex.begin()+4), nullptr, 16));
            b = static_cast<unsigned>(std::stoul(std::string(hex.begin()+4, hex.end()), nullptr, 16));
        }
        char e2[48]; snprintf(e2, sizeof(e2), "\\red%u\\green%u\\blue%u;", r, g, b);
        h += e2;
    }
    h += "}\n\\f2\\fs22\\cf1\\pard\\ql\n";
    return h;
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

// ── HTML → CDM converter ─────────────────────────────────────────────────────

static std::u32string WToU32(const std::wstring& ws) {
    std::u32string s; s.reserve(ws.size());
    for (wchar_t c : ws) s += static_cast<char32_t>(static_cast<uint16_t>(c));
    return s;
}

static cdm::Color HexToColor(const std::wstring& hex) {
    if (hex.size() != 6) return {};
    try {
        auto r = static_cast<uint8_t>(std::stoul(std::string(hex.begin(), hex.begin()+2), nullptr, 16));
        auto g = static_cast<uint8_t>(std::stoul(std::string(hex.begin()+2, hex.begin()+4), nullptr, 16));
        auto b2= static_cast<uint8_t>(std::stoul(std::string(hex.begin()+4, hex.end()), nullptr, 16));
        return cdm::Color::Make(r, g, b2);
    } catch (...) { return {}; }
}

static cdm::Document HtmlToCdm(const std::wstring& html) {
    cdm::DocumentBuilder b;
    b.SetOriginalFormat(cdm::FileFormat::HTML);
    b.BeginSection();

    bool inScript = false, inStyle = false;
    bool paraOpen = false;
    bool inTable = false;
    bool listOpen = false, listOl = false;
    int  listNum = 0;

    // Buffer modes for contexts that only accept plain text
    enum class BufferMode { None, Heading, ListItem, TableCell };
    BufferMode bufMode = BufferMode::None;
    int  headingLevel = 0;
    std::wstring textBuf;

    // Inline style accumulators
    int boldD = 0, italicD = 0, underD = 0, codeD = 0;
    struct ColorEntry { cdm::Color c; bool valid; };
    std::vector<ColorEntry> colorStack;

    auto currentStyle = [&]() {
        cdm::TextStyle s;
        if (boldD > 0)   s.bold   = true;
        if (italicD > 0) s.italic = true;
        if (underD > 0)  s.underline = cdm::UnderlineStyle::Single;
        if (codeD > 0)   s.fontFamily = std::string("Courier New");
        for (auto it = colorStack.rbegin(); it != colorStack.rend(); ++it)
            if (it->valid) { s.color = it->c; break; }
        return s;
    };

    auto hasStyle = [&]() {
        return boldD > 0 || italicD > 0 || underD > 0 || codeD > 0 ||
               (!colorStack.empty() && colorStack.back().valid);
    };

    auto ensureParaOpen = [&]() {
        if (!paraOpen && bufMode == BufferMode::None) {
            b.BeginParagraph(); paraOpen = true;
        }
    };

    auto closePara = [&]() {
        if (paraOpen) { b.EndParagraph(); paraOpen = false; }
    };

    auto flushBuffer = [&]() {
        switch (bufMode) {
            case BufferMode::Heading:
                b.AddHeading(headingLevel, WToU32(textBuf));
                break;
            case BufferMode::ListItem:
                if (!listOpen) {
                    b.BeginList(listOl ? cdm::ListType::Numbered : cdm::ListType::Bullet,
                                listNum > 0 ? listNum : 1);
                    listOpen = true;
                }
                b.AddListItem(WToU32(textBuf));
                break;
            case BufferMode::TableCell:
                b.AddTableCell(WToU32(textBuf));
                break;
            default: break;
        }
        textBuf.clear();
        bufMode = BufferMode::None;
    };

    auto closeList = [&]() {
        if (listOpen) { b.EndList(); listOpen = false; }
    };

    auto addTextContent = [&](const std::wstring& text) {
        if (text.empty()) return;
        if (bufMode != BufferMode::None) {
            textBuf += text;
            return;
        }
        ensureParaOpen();
        auto st = currentStyle();
        auto u32 = WToU32(text);
        if (hasStyle()) b.AddStyledText(u32, st);
        else            b.AddText(u32);
    };

    size_t pos = 0;
    while (pos < html.size()) {
        if (html[pos] != L'<') {
            if (!inScript && !inStyle) {
                size_t end = html.find(L'<', pos);
                if (end == std::wstring::npos) end = html.size();
                std::wstring raw = DecodeEntities(html.substr(pos, end - pos));
                std::wstring norm;
                bool lastWs = false;
                for (wchar_t c : raw) {
                    if (iswspace(c) && c != L'\n') {
                        if (!lastWs && !norm.empty()) { norm += L' '; lastWs = true; }
                    } else if (c != L'\n') {
                        norm += c; lastWs = false;
                    }
                }
                addTextContent(norm);
                pos = end;
            } else {
                ++pos;
            }
            continue;
        }

        size_t lt = pos;
        size_t gt = html.find(L'>', lt);
        if (gt == std::wstring::npos) { ++pos; continue; }

        std::wstring raw = html.substr(lt + 1, gt - lt - 1);
        pos = gt + 1;
        if (raw.empty()) continue;

        if (raw.size() >= 3 && raw.substr(0, 3) == L"!--") {
            size_t ce = html.find(L"-->", lt);
            if (ce != std::wstring::npos) pos = ce + 3;
            continue;
        }

        bool isE = (!raw.empty() && raw[0] == L'/');
        std::wstring nm;
        { size_t i = isE ? 1 : 0;
          while (i < raw.size() && raw[i] != L' ' && raw[i] != L'/' && raw[i] != L'\t')
              nm += towlower(raw[i++]); }

        if (nm == L"script") { inScript = !isE; continue; }
        if (nm == L"style")  { inStyle  = !isE; continue; }
        if (inScript || inStyle) continue;

        if (!isE) {
            if (nm>=L"h1" && nm<=L"h6" && nm.size()==2 && nm[0]==L'h') {
                closePara(); flushBuffer(); closeList();
                headingLevel = nm[1] - L'0';
                bufMode = BufferMode::Heading; textBuf.clear();
            }
            else if (nm==L"p"||nm==L"div") { closePara(); flushBuffer(); closeList(); }
            else if (nm==L"br") {
                if (bufMode != BufferMode::None) textBuf += L'\n';
                else if (paraOpen) b.AddLineBreak();
            }
            else if (nm==L"hr") { closePara(); b.AddHorizontalRule(); }
            else if (nm==L"ul") { closePara(); flushBuffer(); listOl = false; }
            else if (nm==L"ol") { closePara(); flushBuffer(); listOl = true; listNum = 0; }
            else if (nm==L"li") {
                flushBuffer();
                bufMode = BufferMode::ListItem; textBuf.clear(); listNum++;
            }
            else if (nm==L"pre"||nm==L"code"||nm==L"tt") { codeD++; }
            else if (nm==L"b"||nm==L"strong") { boldD++;   }
            else if (nm==L"i"||nm==L"em")     { italicD++; }
            else if (nm==L"u"||nm==L"a")      { underD++;  }
            else if (nm==L"span"||nm==L"font") {
                std::wstring hex;
                if (nm == L"font") { auto cv = HtmlAttr(raw, L"color"); if (!cv.empty()) hex = ParseCssColor(cv); }
                if (hex.empty())   { auto sv = HtmlAttr(raw, L"style"); if (!sv.empty()) hex = ParseCssColor(CssProp(sv, L"color")); }
                if (!hex.empty()) colorStack.push_back({HexToColor(hex), true});
                else              colorStack.push_back({{}, false});
            }
            else if (nm==L"table") { closePara(); flushBuffer(); closeList(); b.BeginTable(); inTable = true; }
            else if (nm==L"tr")    { b.BeginTableRow(); }
            else if (nm==L"td")    { bufMode = BufferMode::TableCell; textBuf.clear(); }
            else if (nm==L"th")    { bufMode = BufferMode::TableCell; textBuf.clear(); }
            else if (nm==L"blockquote") { closePara(); b.BeginBlockQuote(); }
        } else {
            if (nm.size()==2 && nm[0]==L'h' && nm[1]>=L'1' && nm[1]<=L'6') {
                flushBuffer();
            }
            else if (nm==L"p"||nm==L"div") { closePara(); }
            else if (nm==L"li")  { flushBuffer(); }
            else if (nm==L"ul"||nm==L"ol") { flushBuffer(); closeList(); }
            else if (nm==L"pre"||nm==L"code"||nm==L"tt") { codeD = std::max(0, codeD-1); }
            else if (nm==L"b"||nm==L"strong") { boldD   = std::max(0, boldD-1);   }
            else if (nm==L"i"||nm==L"em")     { italicD = std::max(0, italicD-1); }
            else if (nm==L"u"||nm==L"a")      { underD  = std::max(0, underD-1);  }
            else if (nm==L"span"||nm==L"font") { if (!colorStack.empty()) colorStack.pop_back(); }
            else if (nm==L"table") { b.EndTable(); inTable = false; }
            else if (nm==L"tr")    { b.EndTableRow(); }
            else if (nm==L"td"||nm==L"th") { flushBuffer(); }
            else if (nm==L"blockquote") { closePara(); b.EndBlockQuote(); }
        }
    }

    closePara();
    flushBuffer();
    closeList();
    b.EndSection();
    return b.MoveBuild();
}

// ── HTML → RTF converter (kept for reference / fallback) ─────────────────────

static std::string HtmlToRtf(const std::wstring& html) {
    std::string body;

    bool inScript = false, inStyle = false;
    int  boldD = 0, italicD = 0, underD = 0, codeD = 0;
    bool listOl = false;
    int  listNum = 0;

    // Color stack: each entry is a colortbl index (0 = no explicit color = use \cf1 default)
    std::vector<int> colorStack;
    std::map<std::wstring, int> colorMap; // hex -> colortbl index 2+

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
            else if (nm==L"span"||nm==L"font") {
                // Check for color attribute or style
                std::wstring hex;
                if (nm == L"font") {
                    auto cv = HtmlAttr(raw, L"color");
                    if (!cv.empty()) hex = ParseCssColor(cv);
                }
                if (hex.empty()) {
                    auto sv = HtmlAttr(raw, L"style");
                    if (!sv.empty()) hex = ParseCssColor(CssProp(sv, L"color"));
                }
                if (!hex.empty()) {
                    auto it = colorMap.find(hex);
                    if (it == colorMap.end()) {
                        int idx = static_cast<int>(colorMap.size()) + 2;
                        colorMap[hex] = idx;
                        it = colorMap.find(hex);
                    }
                    char rtfClr[24];
                    snprintf(rtfClr, sizeof(rtfClr), "{\\cf%d ", it->second);
                    (inCell?cellBuf:body) += rtfClr;
                    colorStack.push_back(it->second);
                } else {
                    colorStack.push_back(0); // no color, just balance the stack
                }
            }
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
            else if (nm==L"span"||nm==L"font") {
                if (!colorStack.empty()) {
                    if (colorStack.back() != 0)
                        (inCell?cellBuf:body) += "}";
                    colorStack.pop_back();
                }
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

    return BuildHtmlRtfHeader(colorMap) + body + "}";
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
    std::wstring html = ReadFileAsText(path);
    if (html.empty()) { r.error = L"Cannot read HTML file."; return r; }

    r.cdmDoc = HtmlToCdm(html);
    r.ok     = true;
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
