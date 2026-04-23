#include "DocxFormat.h"
#include "ZipReader.h"
#include <sstream>
#include <algorithm>

// ── Shared RTF helpers ────────────────────────────────────────────────────────

// Encode a wide string as RTF: escape \{} and use \uN? for non-ASCII
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

// ── XML helpers ───────────────────────────────────────────────────────────────

static std::wstring WAttr(const std::wstring& tag, const wchar_t* name) {
    std::wstring n = std::wstring(name) + L"=\"";
    auto p = tag.find(n);
    if (p == std::wstring::npos) return {};
    p += n.size();
    auto e = tag.find(L'"', p);
    return e != std::wstring::npos ? tag.substr(p, e - p) : std::wstring{};
}

static void DecodeXmlEntities(std::wstring& s) {
    auto replace = [&](const wchar_t* from, const wchar_t* to) {
        std::wstring f(from), t(to);
        size_t pos = 0;
        while ((pos = s.find(f, pos)) != std::wstring::npos) {
            s.replace(pos, f.size(), t); pos += t.size();
        }
    };
    replace(L"&amp;",  L"&");
    replace(L"&lt;",   L"<");
    replace(L"&gt;",   L">");
    replace(L"&quot;", L"\"");
    replace(L"&apos;", L"'");
}

static std::wstring XmlTag(const std::wstring& xml, const std::wstring& tag) {
    std::wstring open  = L"<" + tag + L">";
    std::wstring close = L"</" + tag + L">";
    auto s = xml.find(open);
    if (s == std::wstring::npos) return {};
    s += open.size();
    auto e = xml.find(close, s);
    if (e == std::wstring::npos) return {};
    return xml.substr(s, e - s);
}

static std::string XmlEscape(const std::wstring& text) {
    std::string out;
    int needed = WideCharToMultiByte(CP_UTF8, 0, text.c_str(),
        static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    std::string utf8(needed, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(),
        static_cast<int>(text.size()), utf8.data(), needed, nullptr, nullptr);
    for (char c : utf8) {
        switch (c) {
            case '&': out += "&amp;";  break;
            case '<': out += "&lt;";   break;
            case '>': out += "&gt;";   break;
            case '"': out += "&quot;"; break;
            default:  out += c;
        }
    }
    return out;
}

// ── DOCX → RTF converter ──────────────────────────────────────────────────────

std::string DocxFormat::ParseDocumentXmlToRtf(const std::wstring& xml) {
    std::string body;

    // --- State ---
    bool inPPr  = false, inRPr = false;
    bool inPara = false, inRun = false;
    bool inTable = false, inRow = false, inCell = false;
    bool inDel   = false;  // inside w:del (tracked deletion — skip)

    // Paragraph properties
    int paraAlign = 0;    // 0=left 1=center 2=right 3=justify
    int paraIndLi = 0;    // left indent in twips
    int headingLv = 0;    // 0=none, 1-6

    // Run properties
    bool runB = false, runI = false, runU = false, runS = false;
    int  runFsz = 0;  // half-points; 0 = inherit

    // Table
    struct Cell { int w; std::string c; };
    std::vector<std::vector<Cell>> tblRows;
    std::vector<Cell>              tblRow;
    std::string                    cellBuf;
    int                            cellW = 0;

    auto& destRef = [&]() -> std::string& {  // helper not used directly; use lambda below
        return inCell ? cellBuf : body;
    };
    (void)destRef;

    size_t pos = 0;
    while (pos < xml.size()) {
        size_t lt = xml.find(L'<', pos);
        if (lt == std::wstring::npos) break;
        size_t gt = xml.find(L'>', lt);
        if (gt == std::wstring::npos) break;

        const std::wstring raw = xml.substr(lt + 1, gt - lt - 1);
        pos = gt + 1;

        if (raw.empty() || raw[0] == L'?' || raw[0] == L'!') continue;

        const bool isE  = (raw[0] == L'/');
        const bool isSC = (raw.back() == L'/');

        std::wstring nm;
        { size_t i = isE ? 1 : 0;
          while (i < raw.size() && raw[i] != L' ' && raw[i] != L'/' &&
                 raw[i] != L'\t' && raw[i] != L'\n') nm += raw[i++]; }

        std::string& dest = inCell ? cellBuf : body;

        // --- Track deletion blocks (skip their content) ---
        if (nm == L"w:del")  { inDel = !isE;  continue; }
        if (inDel)           continue;

        // --- Paragraph ---
        if (nm == L"w:p") {
            if (!isE && !isSC) {
                inPara = true;
                paraAlign = 0; paraIndLi = 0; headingLv = 0;
                if (!inCell) body += "\\pard\\ql ";
            } else if (isE) {
                if (headingLv > 0 && !inCell) body += "\\b0\\fs22 ";
                (inCell ? cellBuf : body) += "\\line ";
                inPara = false;
            }
        }
        else if (nm == L"w:pPr") {
            if (!isE) { inPPr = true; }
            else {
                inPPr = false;
                if (!inCell) {
                    static const char* aln[] = {"\\ql","\\qc","\\qr","\\qj"};
                    body += "\\pard";
                    body += aln[std::max(0, std::min(3, paraAlign))];
                    if (paraIndLi > 0) { body += "\\li"; body += std::to_string(paraIndLi); }
                    if (headingLv > 0) {
                        static const int sz[] = {0,36,28,24,22,20,18};
                        body += "\\b\\fs";
                        body += std::to_string(sz[std::min(6, headingLv)]);
                    }
                    body += " ";
                }
            }
        }
        else if (nm == L"w:jc" && inPPr) {
            auto v = WAttr(raw, L"w:val");
            paraAlign = (v==L"center")?1:(v==L"right")?2:(v==L"both")?3:0;
        }
        else if (nm == L"w:pStyle" && inPPr) {
            auto v = WAttr(raw, L"w:val");
            for (int i = 1; i <= 6; i++) {
                if (v == L"Heading"  + std::to_wstring(i) ||
                    v == L"heading"  + std::to_wstring(i) ||
                    v == L"Heading " + std::to_wstring(i))
                { headingLv = i; break; }
            }
        }
        else if (nm == L"w:ind" && inPPr) {
            auto v = WAttr(raw, L"w:left");
            if (!v.empty()) paraIndLi = _wtoi(v.c_str());
        }
        // --- Run ---
        else if (nm == L"w:r") {
            if (!isE && !isSC) {
                inRun = true; runB=runI=runU=runS=false; runFsz=0;
            } else inRun = false;
        }
        else if (nm == L"w:rPr") { inRPr = !isE; }
        else if ((nm==L"w:b"  ||nm==L"w:bCs") && inRPr && !isE) { runB = WAttr(raw,L"w:val")!=L"0"; }
        else if ((nm==L"w:i"  ||nm==L"w:iCs") && inRPr && !isE) { runI = WAttr(raw,L"w:val")!=L"0"; }
        else if (nm==L"w:u"   && inRPr && !isE) {
            auto v=WAttr(raw,L"w:val"); runU=v.empty()||(v!=L"none"&&v!=L"0");
        }
        else if (nm==L"w:strike" && inRPr && !isE) { runS = WAttr(raw,L"w:val")!=L"0"; }
        else if ((nm==L"w:sz"||nm==L"w:szCs") && inRPr && !isE && runFsz==0) {
            auto v=WAttr(raw,L"w:val"); if(!v.empty()) runFsz=_wtoi(v.c_str());
        }
        // --- Text ---
        else if (nm == L"w:t" && !isE) {
            size_t ce = xml.find(L"</w:t>", gt + 1);
            if (ce != std::wstring::npos) {
                std::wstring text = xml.substr(gt + 1, ce - gt - 1);
                DecodeXmlEntities(text);
                std::string& d2 = inCell ? cellBuf : body;
                bool hasFmt = runB||runI||runU||runS||runFsz>0;
                if (hasFmt) {
                    std::string fmt = "{";
                    if (runB)    fmt += "\\b ";
                    if (runI)    fmt += "\\i ";
                    if (runU)    fmt += "\\ul ";
                    if (runS)    fmt += "\\strike ";
                    if (runFsz > 0) { fmt += "\\fs"; fmt += std::to_string(runFsz); fmt += " "; }
                    d2 += fmt + RtfEnc(text) + "}";
                } else {
                    d2 += RtfEnc(text);
                }
                pos = ce + 6;
            }
        }
        else if (nm == L"w:br" && !isE) {
            auto v = WAttr(raw, L"w:type");
            (inCell ? cellBuf : body) += (v==L"page") ? "\\page " : "\\line ";
        }
        else if (nm == L"w:tab" && !isE) {
            (inCell ? cellBuf : body) += "\\tab ";
        }
        // --- Table ---
        else if (nm == L"w:tbl") {
            if (!isE && !isSC) { inTable=true; tblRows.clear(); }
            else if (isE) {
                if (!tblRows.empty()) {
                    int maxC = 0;
                    for (auto& r : tblRows) maxC = std::max(maxC, (int)r.size());
                    int defW = maxC ? 9360/maxC : 9360;

                    for (auto& row : tblRows) {
                        body += "\\trowd\\trqc\n";
                        int x = 0;
                        for (auto& cl : row) {
                            x += (cl.w > 0 ? cl.w : defW);
                            body += "\\clbrdrt\\brdrw15\\brdrs"
                                    "\\clbrdrl\\brdrw15\\brdrs"
                                    "\\clbrdrb\\brdrw15\\brdrs"
                                    "\\clbrdrr\\brdrw15\\brdrs";
                            body += "\\cellx"; body += std::to_string(x); body += "\n";
                        }
                        for (auto& cl : row) {
                            // trim trailing \line from cell content
                            std::string c = cl.c;
                            while (c.size() >= 6 &&
                                   c.substr(c.size()-6) == "\\line ")
                                c.resize(c.size()-6);
                            body += "\\pard\\intbl "; body += c; body += "\\cell\n";
                        }
                        body += "\\row\n";
                    }
                    body += "\\pard\\ql\n";
                }
                inTable = false;
            }
        }
        else if (nm == L"w:tr") {
            if (!isE && !isSC) { inRow=true; tblRow.clear(); }
            else if (isE)      { tblRows.push_back(tblRow); inRow=false; }
        }
        else if (nm == L"w:tc") {
            if (!isE && !isSC) { inCell=true; cellBuf.clear(); cellW=0; }
            else if (isE)      { tblRow.push_back({cellW, cellBuf}); inCell=false; }
        }
        else if (nm == L"w:tcW" && inCell && !isE) {
            auto w = WAttr(raw, L"w:w");
            auto t = WAttr(raw, L"w:type");
            if (!w.empty() && t != L"pct" && t != L"nil") {
                if (t == L"pct")
                    cellW = (_wtoi(w.c_str()) * 9360) / 5000;
                else
                    cellW = _wtoi(w.c_str());
            }
        }
    }

    return kRtfHeader + body + "}";
}

// ── Build helpers (Save) ──────────────────────────────────────────────────────

std::string DocxFormat::BuildDocumentXml(const std::wstring& text,
                                          const DocProperties& /*props*/) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        << "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n"
        << "<w:body>\r\n";
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        xml << "<w:p><w:r><w:t xml:space=\"preserve\">"
            << XmlEscape(line) << "</w:t></w:r></w:p>\r\n";
    }
    xml << "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
        << "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>"
        << "</w:sectPr>\r\n</w:body>\r\n</w:document>\r\n";
    return xml.str();
}

std::string DocxFormat::BuildCoreXml(const DocProperties& props) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        << "<cp:coreProperties "
        << "xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
        << "xmlns:dc=\"http://purl.org/dc/elements/1.1/\">\r\n"
        << "<dc:title>" << XmlEscape(props.title) << "</dc:title>\r\n"
        << "<dc:creator>" << XmlEscape(props.author) << "</dc:creator>\r\n"
        << "</cp:coreProperties>\r\n";
    return xml.str();
}

std::string DocxFormat::BuildContentTypes() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n"
           "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n"
           "<Default Extension=\"xml\"  ContentType=\"application/xml\"/>\r\n"
           "<Override PartName=\"/word/document.xml\" "
           "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\r\n"
           "<Override PartName=\"/word/styles.xml\" "
           "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\r\n"
           "<Override PartName=\"/word/settings.xml\" "
           "ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>\r\n"
           "<Override PartName=\"/docProps/core.xml\" "
           "ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\r\n"
           "</Types>\r\n";
}

std::string DocxFormat::BuildRels() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"
           "<Relationship Id=\"rId1\" "
           "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
           "Target=\"word/document.xml\"/>\r\n"
           "<Relationship Id=\"rId2\" "
           "Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" "
           "Target=\"docProps/core.xml\"/>\r\n"
           "</Relationships>\r\n";
}

std::string DocxFormat::BuildWordRels() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"
           "<Relationship Id=\"rId1\" "
           "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
           "Target=\"styles.xml\"/>\r\n"
           "<Relationship Id=\"rId2\" "
           "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" "
           "Target=\"settings.xml\"/>\r\n"
           "</Relationships>\r\n";
}

std::string DocxFormat::BuildSettings() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n"
           "<w:defaultTabStop w:val=\"720\"/>\r\n"
           "</w:settings>\r\n";
}

std::string DocxFormat::BuildStyles() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\r\n"
           "<w:docDefaults><w:rPrDefault><w:rPr>"
           "<w:rFonts w:ascii=\"Malgun Gothic\" w:eastAsia=\"Malgun Gothic\" w:hAnsi=\"Calibri\"/>"
           "<w:sz w:val=\"22\"/>"
           "</w:rPr></w:rPrDefault></w:docDefaults>\r\n"
           "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">"
           "<w:name w:val=\"Normal\"/></w:style>\r\n"
           "</w:styles>\r\n";
}

// ── Public interface ───────────────────────────────────────────────────────────

FormatResult DocxFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    ZipReader zip;
    if (!zip.Open(path)) {
        r.error = L"Cannot open DOCX file (invalid ZIP).";
        return r;
    }

    if (zip.HasEntry("docProps/core.xml")) {
        std::wstring coreXml = zip.ExtractWide("docProps/core.xml");
        doc.Properties().title   = XmlTag(coreXml, L"dc:title");
        doc.Properties().author  = XmlTag(coreXml, L"dc:creator");
        for (auto* p : { &doc.Properties().title, &doc.Properties().author })
            DecodeXmlEntities(*p);
    }

    if (!zip.HasEntry("word/document.xml")) {
        r.error = L"Malformed DOCX: missing word/document.xml";
        return r;
    }
    std::wstring docXml = zip.ExtractWide("word/document.xml");
    std::string  rtf    = ParseDocumentXmlToRtf(docXml);

    r.content = std::wstring(rtf.begin(), rtf.end());
    r.rtf     = true;
    r.ok      = true;
    return r;
}

FormatResult DocxFormat::Save(const std::wstring& path,
                               const std::wstring& content,
                               const std::string&  /*rtf*/,
                               Document&           doc) {
    FormatResult r;
    ZipWriter zip;
    if (!zip.Open(path)) { r.error = L"Cannot create DOCX file."; return r; }

    const auto& props = doc.Properties();
    zip.AddText("[Content_Types].xml",         BuildContentTypes(), false);
    zip.AddText("_rels/.rels",                 BuildRels(),         false);
    zip.AddText("word/_rels/document.xml.rels",BuildWordRels(),     false);
    zip.AddText("word/document.xml",           BuildDocumentXml(content, props));
    zip.AddText("word/styles.xml",             BuildStyles());
    zip.AddText("word/settings.xml",           BuildSettings());
    zip.AddText("docProps/core.xml",           BuildCoreXml(props));
    zip.Close();
    r.ok = true;
    return r;
}
