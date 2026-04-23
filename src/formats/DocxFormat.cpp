#include "DocxFormat.h"
#include "ZipReader.h"
#include <sstream>
#include <algorithm>

// ---- XML text extraction helpers ----

// Extract text content of a specific XML element (first occurrence)
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

// Decode basic XML entities in place
static void DecodeXmlEntities(std::wstring& s) {
    auto replace = [&](const wchar_t* from, const wchar_t* to) {
        std::wstring f(from), t(to);
        size_t pos = 0;
        while ((pos = s.find(f, pos)) != std::wstring::npos) {
            s.replace(pos, f.size(), t);
            pos += t.size();
        }
    };
    replace(L"&amp;",  L"&");
    replace(L"&lt;",   L"<");
    replace(L"&gt;",   L">");
    replace(L"&quot;", L"\"");
    replace(L"&apos;", L"'");
}

// Escape text for XML output
static std::string XmlEscape(const std::wstring& text) {
    std::string out;
    int needed = WideCharToMultiByte(CP_UTF8, 0, text.c_str(),
                                     static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    std::string utf8(needed, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(),
                        static_cast<int>(text.size()), utf8.data(), needed, nullptr, nullptr);
    for (char c : utf8) {
        switch (c) {
            case '&':  out += "&amp;";  break;
            case '<':  out += "&lt;";   break;
            case '>':  out += "&gt;";   break;
            case '"':  out += "&quot;"; break;
            case '\'': out += "&apos;"; break;
            default:   out += c;
        }
    }
    return out;
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return {};
    int needed = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), static_cast<int>(s.size()), nullptr, 0);
    std::wstring ws(needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), static_cast<int>(s.size()), ws.data(), needed);
    return ws;
}

// ---- Parse word/document.xml ----
std::wstring DocxFormat::ParseDocumentXml(const std::wstring& xml) {
    std::wostringstream out;
    size_t pos = 0;
    bool afterParagraph = false;

    while (pos < xml.size()) {
        // Find next tag
        size_t tagStart = xml.find(L'<', pos);
        if (tagStart == std::wstring::npos) break;
        size_t tagEnd = xml.find(L'>', tagStart);
        if (tagEnd == std::wstring::npos) break;

        std::wstring tag = xml.substr(tagStart + 1, tagEnd - tagStart - 1);
        // Strip attributes from tag name
        size_t sp = tag.find_first_of(L" \t\r\n/");
        std::wstring tagName = (sp != std::wstring::npos) ? tag.substr(0, sp) : tag;

        if (tagName == L"w:p" || tagName == L"w:p>") {
            // Paragraph end: add newline
            if (afterParagraph) out << L'\n';
            afterParagraph = true;
        } else if (tagName == L"w:t") {
            // Text run: extract content
            size_t contentStart = tagEnd + 1;
            size_t contentEnd   = xml.find(L"</w:t>", contentStart);
            if (contentEnd == std::wstring::npos) { pos = tagEnd + 1; continue; }
            std::wstring text = xml.substr(contentStart, contentEnd - contentStart);
            DecodeXmlEntities(text);
            out << text;
            pos = contentEnd + 6;
            continue;
        } else if (tagName == L"w:br") {
            out << L'\n';
        } else if (tagName == L"w:tab") {
            out << L'\t';
        }
        pos = tagEnd + 1;
    }
    return out.str();
}

// ---- Build DOCX XML files ----

std::string DocxFormat::BuildDocumentXml(const std::wstring& text, const DocProperties& /*props*/) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
        << "<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" "
        << "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n"
        << "<w:body>\r\n";

    // Split text into paragraphs by newline
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        xml << "<w:p><w:r><w:t xml:space=\"preserve\">"
            << XmlEscape(line)
            << "</w:t></w:r></w:p>\r\n";
    }

    xml << "<w:sectPr>"
        << "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>"  // Letter
        << "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>"
        << "</w:sectPr>\r\n"
        << "</w:body>\r\n</w:document>\r\n";
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
        << "<dc:subject>" << XmlEscape(props.subject) << "</dc:subject>\r\n"
        << "<cp:keywords>" << XmlEscape(props.keywords) << "</cp:keywords>\r\n"
        << "<dc:description>" << XmlEscape(props.comment) << "</dc:description>\r\n"
        << "</cp:coreProperties>\r\n";
    return xml.str();
}

std::string DocxFormat::BuildContentTypes() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n"
           "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n"
           "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n"
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
           "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>\r\n"
           "</w:settings>\r\n";
}

std::string DocxFormat::BuildStyles() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
           "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
           "w:docDefaults=\"\">\r\n"
           "<w:docDefaults>\r\n"
           "<w:rPrDefault><w:rPr>"
           "<w:rFonts w:ascii=\"맑은 고딕\" w:eastAsia=\"맑은 고딕\" w:hAnsi=\"Calibri\"/>"
           "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/>"
           "</w:rPr></w:rPrDefault>\r\n"
           "</w:docDefaults>\r\n"
           "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">\r\n"
           "<w:name w:val=\"Normal\"/><w:qFormat/>\r\n"
           "</w:style>\r\n"
           "</w:styles>\r\n";
}

// ---- Public interface ----

FormatResult DocxFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    ZipReader zip;
    if (!zip.Open(path)) {
        r.error = L"Cannot open DOCX file (invalid ZIP).";
        return r;
    }

    // Read core properties if available
    if (zip.HasEntry("docProps/core.xml")) {
        std::wstring coreXml = zip.ExtractWide("docProps/core.xml");
        doc.Properties().title   = XmlTag(coreXml, L"dc:title");
        doc.Properties().author  = XmlTag(coreXml, L"dc:creator");
        doc.Properties().subject = XmlTag(coreXml, L"dc:subject");
        doc.Properties().keywords= XmlTag(coreXml, L"cp:keywords");
        doc.Properties().comment = XmlTag(coreXml, L"dc:description");
        for (auto* p : { &doc.Properties().title, &doc.Properties().author,
                         &doc.Properties().subject, &doc.Properties().keywords,
                         &doc.Properties().comment })
            DecodeXmlEntities(*p);
    }

    // Read main document
    if (!zip.HasEntry("word/document.xml")) {
        r.error = L"Malformed DOCX: missing word/document.xml";
        return r;
    }
    std::wstring docXml = zip.ExtractWide("word/document.xml");
    r.content = ParseDocumentXml(docXml);
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult DocxFormat::Save(const std::wstring& path,
                               const std::wstring& content,
                               const std::string&  /*rtf*/,
                               Document&           doc) {
    FormatResult r;
    ZipWriter zip;
    if (!zip.Open(path)) {
        r.error = L"Cannot create DOCX file.";
        return r;
    }

    const auto& props = doc.Properties();
    zip.AddText("[Content_Types].xml", BuildContentTypes(), false);
    zip.AddText("_rels/.rels",         BuildRels(),         false);
    zip.AddText("word/_rels/document.xml.rels", BuildWordRels(), false);
    zip.AddText("word/document.xml",   BuildDocumentXml(content, props));
    zip.AddText("word/styles.xml",     BuildStyles());
    zip.AddText("word/settings.xml",   BuildSettings());
    zip.AddText("docProps/core.xml",   BuildCoreXml(props));
    zip.Close();

    r.ok = true;
    return r;
}
