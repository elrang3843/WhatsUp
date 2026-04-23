#include "HwpxFormat.h"
#include "ZipReader.h"
#include <sstream>
#include <algorithm>

// Case-insensitive ZIP entry lookup — real HWPX files vary in capitalization
static std::string FindEntryCI(const ZipReader& zip, const std::string& name) {
    std::string lower = name;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    for (auto& [key, _] : zip.Entries()) {
        std::string klower = key;
        std::transform(klower.begin(), klower.end(), klower.begin(), ::tolower);
        if (klower == lower) return key;
    }
    return {};
}

static bool HasEntryCI(const ZipReader& zip, const std::string& name) {
    return !FindEntryCI(zip, name).empty();
}
static std::wstring ExtractWideCI(ZipReader& zip, const std::string& name) {
    std::string real = FindEntryCI(zip, name);
    return real.empty() ? std::wstring{} : zip.ExtractWide(real);
}

static std::wstring XmlTagW(const std::wstring& xml, const std::wstring& tag) {
    auto open  = L"<" + tag + L">";
    auto close = L"</" + tag + L">";
    auto s = xml.find(open);
    if (s == std::wstring::npos) return {};
    s += open.size();
    auto e = xml.find(close, s);
    return (e == std::wstring::npos) ? std::wstring{} : xml.substr(s, e - s);
}

static void DecodeXml(std::wstring& s) {
    auto replace = [&](const wchar_t* f, const wchar_t* t) {
        std::wstring from(f), to(t);
        size_t pos = 0;
        while ((pos = s.find(from, pos)) != std::wstring::npos) {
            s.replace(pos, from.size(), to); pos += to.size();
        }
    };
    replace(L"&amp;", L"&"); replace(L"&lt;", L"<");
    replace(L"&gt;", L">"); replace(L"&quot;", L"\""); replace(L"&apos;", L"'");
}

static std::string XmlEsc(const std::wstring& text) {
    int n = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()),
                                nullptr, 0, nullptr, nullptr);
    std::string utf8(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()),
                        utf8.data(), n, nullptr, nullptr);
    std::string out;
    for (char c : utf8) {
        if      (c == '&')  out += "&amp;";
        else if (c == '<')  out += "&lt;";
        else if (c == '>')  out += "&gt;";
        else                out += c;
    }
    return out;
}

// HWPX BodyText XML: text is inside <hp:t> elements within <hp:p> paragraphs
std::wstring HwpxFormat::ParseBodyText(const std::wstring& xml) {
    std::wostringstream out;
    size_t pos = 0;
    while (pos < xml.size()) {
        size_t tagStart = xml.find(L'<', pos);
        if (tagStart == std::wstring::npos) break;
        size_t tagEnd = xml.find(L'>', tagStart);
        if (tagEnd == std::wstring::npos) break;

        std::wstring tag = xml.substr(tagStart + 1, tagEnd - tagStart - 1);
        size_t sp = tag.find_first_of(L" /\t");
        std::wstring name = (sp != std::wstring::npos) ? tag.substr(0, sp) : tag;

        if (name == L"hp:p" || name == L"hh:p") {
            out << L'\n';
        } else if (name == L"hp:t" || name == L"hh:t") {
            size_t cs = tagEnd + 1;
            std::wstring closeTag = L"</" + name + L">";
            size_t ce = xml.find(closeTag, cs);
            if (ce == std::wstring::npos) { pos = tagEnd + 1; continue; }
            std::wstring text = xml.substr(cs, ce - cs);
            DecodeXml(text);
            out << text;
            pos = ce + closeTag.size();
            continue;
        }
        pos = tagEnd + 1;
    }
    return out.str();
}

std::string HwpxFormat::BuildBodyText(const std::wstring& text) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
        << "<hs:BodyText xmlns:hs=\"http://www.hancom.co.kr/hwpml/2011/section\" "
        << "xmlns:hp=\"http://www.hancom.co.kr/hwpml/2011/paragraph\">\r\n"
        << "<hs:SectionList><hs:Section>\r\n";

    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        xml << "<hp:P><hp:Run><hp:T><![CDATA[" << XmlEsc(line) << "]]></hp:T></hp:Run></hp:P>\r\n";
    }

    xml << "</hs:Section></hs:SectionList>\r\n</hs:BodyText>\r\n";
    return xml.str();
}

std::string HwpxFormat::BuildContentType() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
           "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n"
           "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n"
           "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n"
           "<Override PartName=\"/Contents/content.hpf\" ContentType=\"application/hwpml-package+xml\"/>\r\n"
           "<Override PartName=\"/Contents/Sections/Section0.xml\" ContentType=\"application/hwpml-section+xml\"/>\r\n"
           "</Types>\r\n";
}

std::string HwpxFormat::BuildRels() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
           "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"
           "<Relationship Id=\"rId1\" Type=\"http://www.hancom.co.kr/hwpml/2011/package/relationships/hwp-package\" Target=\"Contents/content.hpf\"/>\r\n"
           "</Relationships>\r\n";
}

std::string HwpxFormat::BuildManifest() {
    return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
           "<hpf:HWPFPackage xmlns:hpf=\"http://www.hancom.co.kr/hwpml/2011/package\">\r\n"
           "<hpf:Metadata><hpf:Application>WhatsUp 1.0</hpf:Application></hpf:Metadata>\r\n"
           "<hpf:Manifest>\r\n"
           "<hpf:Item Id=\"Sections/Section0\" MediaType=\"application/hwpml-section+xml\" href=\"Sections/Section0.xml\"/>\r\n"
           "</hpf:Manifest>\r\n"
           "</hpf:HWPFPackage>\r\n";
}

std::string HwpxFormat::BuildSummary(const DocProperties& props) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n"
        << "<hsp:HWPSummaryInfo xmlns:hsp=\"http://www.hancom.co.kr/hwpml/2011/summaryinfo\">\r\n"
        << "<hsp:Title>" << XmlEsc(props.title) << "</hsp:Title>\r\n"
        << "<hsp:Author>" << XmlEsc(props.author) << "</hsp:Author>\r\n"
        << "<hsp:Subject>" << XmlEsc(props.subject) << "</hsp:Subject>\r\n"
        << "<hsp:Description>" << XmlEsc(props.comment) << "</hsp:Description>\r\n"
        << "</hsp:HWPSummaryInfo>\r\n";
    return xml.str();
}

FormatResult HwpxFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    ZipReader zip;
    if (!zip.Open(path)) {
        r.error = L"Cannot open HWPX file.";
        return r;
    }

    // Try reading summary info (case-insensitive)
    for (const char* s : { "Contents/summary.xml", "Summary.xml", "Contents/Summary.xml" }) {
        if (HasEntryCI(zip, s)) {
            std::wstring sumXml = ExtractWideCI(zip, s);
            doc.Properties().title   = XmlTagW(sumXml, L"hsp:Title");
            doc.Properties().author  = XmlTagW(sumXml, L"hsp:Author");
            doc.Properties().subject = XmlTagW(sumXml, L"hsp:Subject");
            doc.Properties().comment = XmlTagW(sumXml, L"hsp:Description");
            for (auto* p : {&doc.Properties().title, &doc.Properties().author,
                             &doc.Properties().subject, &doc.Properties().comment})
                DecodeXml(*p);
            break;
        }
    }

    // Find section files — try several path patterns with case-insensitive matching
    std::wostringstream combined;
    for (int i = 0; i < 100; ++i) {
        std::string found;
        // Try common path variants for section files
        for (const char* fmt : {
                "Contents/Sections/Section%d.xml",
                "Contents/sections/section%d.xml",
                "BodyText/Section%d.xml",
                "bodytext/section%d.xml",
                "Contents/section%d.xml" }) {
            char buf[64];
            snprintf(buf, sizeof(buf), fmt, i);
            std::string real = FindEntryCI(zip, buf);
            if (!real.empty()) { found = real; break; }
        }
        if (found.empty()) break;
        std::wstring sxml = zip.ExtractWide(found);
        combined << ParseBodyText(sxml);
    }

    r.content = combined.str();
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult HwpxFormat::Save(const std::wstring& path,
                               const std::wstring& content,
                               const std::string&  /*rtf*/,
                               Document&           doc) {
    FormatResult r;
    ZipWriter zip;
    if (!zip.Open(path)) { r.error = L"Cannot create HWPX file."; return r; }

    zip.AddText("[Content_Types].xml",          BuildContentType(), false);
    zip.AddText("_rels/.rels",                   BuildRels(),        false);
    zip.AddText("Contents/content.hpf",          BuildManifest());
    zip.AddText("Contents/summary.xml",          BuildSummary(doc.Properties()));
    zip.AddText("Contents/Sections/Section0.xml",BuildBodyText(content));
    zip.Close();

    r.ok = true;
    return r;
}
