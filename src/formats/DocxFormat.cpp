#include "DocxFormat.h"
#include "ZipReader.h"
#include "../cdm/document_builder.hpp"
#include <sstream>
#include <algorithm>
#include <map>

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

// ── DOCX → CDM converter ─────────────────────────────────────────────────────

static std::u32string WToU32(const std::wstring& ws) {
    std::u32string s; s.reserve(ws.size());
    for (wchar_t c : ws) s += static_cast<char32_t>(static_cast<uint16_t>(c));
    return s;
}

static cdm::Color HexToColorDocx(const std::wstring& hex) {
    if (hex.size() != 6) return {};
    try {
        auto r = static_cast<uint8_t>(std::stoul(std::string(hex.begin(),   hex.begin()+2), nullptr, 16));
        auto g = static_cast<uint8_t>(std::stoul(std::string(hex.begin()+2, hex.begin()+4), nullptr, 16));
        auto b = static_cast<uint8_t>(std::stoul(std::string(hex.begin()+4, hex.end()),     nullptr, 16));
        return cdm::Color::Make(r, g, b);
    } catch (...) { return {}; }
}

cdm::Document DocxFormat::ParseDocumentXmlToCdm(const std::wstring& xml) {
    cdm::DocumentBuilder b;
    b.SetOriginalFormat(cdm::FileFormat::DOCX);
    b.BeginSection();

    bool inPPr = false, inRPr = false;
    bool inPara = false, inRun = false;
    bool inDel = false;
    bool paraOpen = false;

    int  paraAlign = 0, headingLv = 0;
    bool runB = false, runI = false, runU = false, runS = false;
    int  runFsz = 0;
    std::wstring runColor;

    bool inTable = false, inRow = false, inCell = false;
    std::wstring cellBuf;

    struct CellEntry { std::wstring text; };
    std::vector<std::vector<CellEntry>> tblRows;
    std::vector<CellEntry>              tblRow;

    auto ensureParaOpen = [&]() {
        if (!paraOpen && !inCell) {
            b.BeginParagraph();
            if (headingLv > 0) {
                // heading styling via Bold+fontSize via TextStyle — applied at text node level
            }
            if (paraAlign == 1) b.SetCurrentParagraphAlignment(cdm::Alignment::Center);
            else if (paraAlign == 2) b.SetCurrentParagraphAlignment(cdm::Alignment::Right);
            else if (paraAlign == 3) b.SetCurrentParagraphAlignment(cdm::Alignment::Justify);
            paraOpen = true;
        }
    };

    auto addRunText = [&](const std::wstring& text) {
        if (inCell) { cellBuf += text; return; }
        ensureParaOpen();
        cdm::TextStyle st;
        if (runB) st.bold = true;
        if (runI) st.italic = true;
        if (runU) st.underline = cdm::UnderlineStyle::Single;
        if (runS) st.strike = true;
        if (runFsz > 0) st.fontSize = cdm::Length::Pt(runFsz / 2.0);
        if (headingLv > 0) {
            static const double hSz[] = {0, 18, 14, 12, 11, 10, 9};
            st.fontSize = cdm::Length::Pt(hSz[std::min(6, headingLv)]);
            st.bold = true;
        }
        if (!runColor.empty()) st.color = HexToColorDocx(runColor);
        auto u32 = WToU32(text);
        bool hasStyle = runB || runI || runU || runS || runFsz > 0 ||
                        !runColor.empty() || headingLv > 0;
        if (hasStyle) b.AddStyledText(u32, st);
        else          b.AddText(u32);
    };

    auto closePara = [&]() {
        if (paraOpen) { b.EndParagraph(); paraOpen = false; }
    };

    auto flushTable = [&]() {
        if (tblRows.empty()) return;
        b.BeginTable();
        for (auto& row : tblRows) {
            b.BeginTableRow();
            for (auto& cell : row)
                b.AddTableCell(WToU32(cell.text));
            b.EndTableRow();
        }
        b.EndTable();
        tblRows.clear();
    };

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

        if (nm == L"w:del") { inDel = !isE; continue; }
        if (inDel) continue;

        if (nm == L"w:p") {
            if (!isE && !isSC) {
                inPara = true;
                paraAlign = 0; headingLv = 0;
            } else if (isE) {
                closePara();
                inPara = false;
            }
        }
        else if (nm == L"w:pPr") { inPPr = !isE; }
        else if (nm == L"w:jc" && inPPr) {
            auto v = WAttr(raw, L"w:val");
            paraAlign = (v==L"center")?1:(v==L"right")?2:(v==L"both")?3:0;
        }
        else if (nm == L"w:pStyle" && inPPr) {
            auto v = WAttr(raw, L"w:val");
            for (int lv = 1; lv <= 6; lv++) {
                if (v == L"Heading"  + std::to_wstring(lv) ||
                    v == L"heading"  + std::to_wstring(lv) ||
                    v == L"Heading " + std::to_wstring(lv))
                { headingLv = lv; break; }
            }
        }
        else if (nm == L"w:r") {
            if (!isE && !isSC) {
                inRun = true; runB=runI=runU=runS=false; runFsz=0; runColor.clear();
            } else inRun = false;
        }
        else if (nm == L"w:rPr") { inRPr = !isE; }
        else if ((nm==L"w:b"||nm==L"w:bCs")  && inRPr && !isE) { runB = WAttr(raw,L"w:val")!=L"0"; }
        else if ((nm==L"w:i"||nm==L"w:iCs")  && inRPr && !isE) { runI = WAttr(raw,L"w:val")!=L"0"; }
        else if (nm==L"w:u"                   && inRPr && !isE) {
            auto v=WAttr(raw,L"w:val"); runU=v.empty()||(v!=L"none"&&v!=L"0");
        }
        else if (nm==L"w:strike"              && inRPr && !isE) { runS = WAttr(raw,L"w:val")!=L"0"; }
        else if ((nm==L"w:sz"||nm==L"w:szCs") && inRPr && !isE && runFsz==0) {
            auto v=WAttr(raw,L"w:val"); if(!v.empty()) runFsz=_wtoi(v.c_str());
        }
        else if (nm==L"w:color" && inRPr && !isE) {
            auto v = WAttr(raw, L"w:val");
            if (v != L"auto" && v.size() == 6) runColor = v;
        }
        else if (nm == L"w:t" && !isE) {
            size_t ce = xml.find(L"</w:t>", gt + 1);
            if (ce != std::wstring::npos) {
                std::wstring text = xml.substr(gt + 1, ce - gt - 1);
                DecodeXmlEntities(text);
                addRunText(text);
                pos = ce + 6;
            }
        }
        else if (nm == L"w:br" && !isE) {
            auto v = WAttr(raw, L"w:type");
            if (inCell) cellBuf += L'\n';
            else if (paraOpen) {
                if (v == L"page") { closePara(); }
                else b.AddLineBreak();
            }
        }
        else if (nm == L"w:tab" && !isE) {
            if (inCell) cellBuf += L'\t';
            else { ensureParaOpen(); b.AddTab(); }
        }
        else if (nm == L"w:tbl") {
            if (!isE && !isSC) { closePara(); inTable=true; tblRows.clear(); }
            else if (isE) { flushTable(); inTable=false; }
        }
        else if (nm == L"w:tr") {
            if (!isE && !isSC) { inRow=true; tblRow.clear(); }
            else if (isE) { tblRows.push_back(tblRow); inRow=false; }
        }
        else if (nm == L"w:tc") {
            if (!isE && !isSC) { inCell=true; cellBuf.clear(); }
            else if (isE) { tblRow.push_back({cellBuf}); inCell=false; }
        }
    }

    closePara();
    b.EndSection();
    return b.MoveBuild();
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
    r.cdmDoc = ParseDocumentXmlToCdm(docXml);
    r.ok     = true;
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
