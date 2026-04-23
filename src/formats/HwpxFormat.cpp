#include "HwpxFormat.h"
#include "ZipReader.h"
#include "../cdm/document_builder.hpp"
#include <sstream>
#include <algorithm>
#include <unordered_map>

// ── ZIP helpers ───────────────────────────────────────────────────────────────

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

// ── XML helpers ───────────────────────────────────────────────────────────────

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
    auto rep = [&](const wchar_t* f, const wchar_t* t) {
        std::wstring from(f), to(t); size_t p = 0;
        while ((p = s.find(from, p)) != std::wstring::npos) { s.replace(p, from.size(), to); p += to.size(); }
    };
    rep(L"&amp;", L"&"); rep(L"&lt;", L"<"); rep(L"&gt;", L">");
    rep(L"&quot;", L"\""); rep(L"&apos;", L"'");
}

static std::string XmlEsc(const std::wstring& text) {
    int n = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), nullptr, 0, nullptr, nullptr);
    std::string utf8(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), utf8.data(), n, nullptr, nullptr);
    std::string out;
    for (char c : utf8) {
        if      (c=='&') out += "&amp;";
        else if (c=='<') out += "&lt;";
        else if (c=='>') out += "&gt;";
        else             out += c;
    }
    return out;
}

// Case-insensitive substring search (used by legacy ParseBodyText)
static size_t FindCI(const std::wstring& s, const std::wstring& pat, size_t from) {
    if (pat.empty()) return from;
    for (size_t i = from; i + pat.size() <= s.size(); i++) {
        bool ok = true;
        for (size_t k = 0; k < pat.size() && ok; k++)
            if (towlower(s[i+k]) != towlower(pat[k])) ok = false;
        if (ok) return i;
    }
    return std::wstring::npos;
}

// ── HWPX parsing infrastructure ───────────────────────────────────────────────

namespace {

struct CharShapeInfo {
    double ptSize    = 0.0;   // 0 = use default
    bool   bold      = false;
    bool   italic    = false;
    bool   underline = false;
    bool   strike    = false;
    std::optional<cdm::Color> color;
    bool HasStyle() const {
        return bold || italic || underline || strike || color.has_value() || ptSize > 0.0;
    }
};

struct ParaShapeInfo {
    cdm::Alignment alignment    = cdm::Alignment::Left;
    int            headingLevel = 0;   // 0=none, 1-6=H1-H6
};

using CharShapeMap = std::unordered_map<int, CharShapeInfo>;
using ParaShapeMap = std::unordered_map<int, ParaShapeInfo>;

// Strip namespace prefix and lowercase: "hp:P" → "p", "hh:CharShape" → "charshape"
static std::wstring NormTag(const std::wstring& raw, bool isClose) {
    std::wstring nm;
    size_t i = isClose ? 1 : 0;
    while (i < raw.size() && raw[i]!=L' ' && raw[i]!=L'/' && raw[i]!=L'\t' && raw[i]!=L'\n')
        nm += raw[i++];
    auto c = nm.rfind(L':');
    if (c != std::wstring::npos) nm = nm.substr(c + 1);
    for (auto& ch : nm) ch = towlower(ch);
    return nm;
}

// Case-insensitive attribute lookup
static std::wstring HwpAttr(const std::wstring& tag, const std::wstring& lname) {
    std::wstring ltag = tag;
    for (auto& c : ltag) c = towlower(c);
    std::wstring pat = lname + L"=\"";
    auto p = ltag.find(pat);
    if (p == std::wstring::npos) return {};
    p += pat.size();
    auto e = tag.find(L'"', p);
    return (e != std::wstring::npos) ? tag.substr(p, e - p) : std::wstring{};
}

// Get trimmed text content immediately after a tag's closing '>'
static std::wstring TagText(const std::wstring& xml, size_t afterGt) {
    size_t lt = xml.find(L'<', afterGt);
    std::wstring s = (lt != std::wstring::npos) ? xml.substr(afterGt, lt - afterGt) : xml.substr(afterGt);
    size_t b = 0;            while (b < s.size() && iswspace(s[b])) b++;
    size_t e = s.size();     while (e > b && iswspace(s[e-1])) e--;
    return s.substr(b, e - b);
}

static std::optional<cdm::Color> ParseHwpColor(const std::wstring& hex) {
    if (hex.size() < 6) return {};
    try {
        auto r = (uint8_t)std::stoul(std::string(hex.begin(),   hex.begin()+2), nullptr, 16);
        auto g = (uint8_t)std::stoul(std::string(hex.begin()+2, hex.begin()+4), nullptr, 16);
        auto b = (uint8_t)std::stoul(std::string(hex.begin()+4, hex.begin()+6), nullptr, 16);
        if (!r && !g && !b) return {};   // skip default black
        return cdm::Color::Make(r, g, b);
    } catch (...) { return {}; }
}

static std::u32string WToU32(const std::wstring& ws) {
    std::u32string s; s.reserve(ws.size());
    for (wchar_t c : ws) s += static_cast<char32_t>(static_cast<uint16_t>(c));
    return s;
}

// Parse HWPX header XML for CharShape and ParaShape definitions
void ParseHwpxHeader(const std::wstring& xml, CharShapeMap& chars, ParaShapeMap& paras) {
    enum State { NONE, IN_CHAR, IN_PARA } state = NONE;
    int curId = -1;
    CharShapeInfo cs{};
    ParaShapeInfo ps{};

    size_t pos = 0;
    while (pos < xml.size()) {
        size_t lt = xml.find(L'<', pos); if (lt == std::wstring::npos) break;
        size_t gt = xml.find(L'>', lt);  if (gt == std::wstring::npos) break;
        const std::wstring raw = xml.substr(lt+1, gt-lt-1);
        pos = gt + 1;
        if (raw.empty() || raw[0]==L'?' || raw[0]==L'!') continue;
        bool isE  = (raw[0] == L'/');
        bool isSC = (!raw.empty() && raw.back() == L'/');
        std::wstring nm = NormTag(raw, isE);

        if (nm == L"charshape") {
            if (!isE && !isSC) {
                auto v = HwpAttr(raw, L"id");
                curId = v.empty() ? -1 : _wtoi(v.c_str());
                cs = {}; state = IN_CHAR;
            } else if (isE && state == IN_CHAR && curId >= 0) {
                chars[curId] = cs; state = NONE; curId = -1;
            }
        } else if (nm == L"parashape") {
            if (!isE && !isSC) {
                auto v = HwpAttr(raw, L"id");
                curId = v.empty() ? -1 : _wtoi(v.c_str());
                ps = {}; state = IN_PARA;
            } else if (isE && state == IN_PARA && curId >= 0) {
                paras[curId] = ps; state = NONE; curId = -1;
            }
        } else if (!isE && state == IN_CHAR) {
            std::wstring val = TagText(xml, gt+1);
            if (nm == L"height" && !val.empty()) {
                // HWP unit: 200 per pt (2200 → 11pt)
                double h = _wtof(val.c_str());
                if (h > 0) { cs.ptSize = h / 200.0; if (cs.ptSize < 4 || cs.ptSize > 144) cs.ptSize = 0; }
            } else if (nm == L"bold")
                cs.bold = (val==L"1"||val==L"true"||val==L"TRUE");
            else if (nm == L"italic")
                cs.italic = (val==L"1"||val==L"true"||val==L"TRUE");
            else if (nm == L"underline")
                cs.underline = !val.empty() && val!=L"NONE" && val!=L"0" && val!=L"false";
            else if (nm==L"strikeout" || nm==L"strikethru" || nm==L"strike")
                cs.strike = !val.empty() && val!=L"NONE" && val!=L"0" && val!=L"false";
            else if (nm==L"textcolor" || nm==L"forecolor" || nm==L"color") {
                if (!val.empty()) cs.color = ParseHwpColor(val);
            } else if (nm == L"rgb") {   // <hh:RGB r="0" g="0" b="0"/>
                auto rv = HwpAttr(raw, L"r"), gv = HwpAttr(raw, L"g"), bv = HwpAttr(raw, L"b");
                if (!rv.empty() && !gv.empty() && !bv.empty()) {
                    int r=_wtoi(rv.c_str()), g=_wtoi(gv.c_str()), b=_wtoi(bv.c_str());
                    if (r||g||b) cs.color = cdm::Color::Make((uint8_t)r, (uint8_t)g, (uint8_t)b);
                }
            }
        } else if (!isE && state == IN_PARA) {
            std::wstring val = TagText(xml, gt+1);
            if (nm==L"justify" || nm==L"alignment" || nm==L"align") {
                if      (val==L"CENTER"||val==L"center") ps.alignment = cdm::Alignment::Center;
                else if (val==L"RIGHT" ||val==L"right")  ps.alignment = cdm::Alignment::Right;
                else if (val==L"BOTH"  ||val==L"both"  ||val==L"DISTRIBUTE") ps.alignment = cdm::Alignment::Justify;
                else ps.alignment = cdm::Alignment::Left;
            } else if (nm==L"outline"||nm==L"headingtype"||nm==L"heading") {
                if (!val.empty() && val!=L"NONE" && val!=L"0") {
                    for (int lv=1; lv<=6; lv++) {
                        if (val==L"H"+std::to_wstring(lv) ||
                            val==L"HEADING"+std::to_wstring(lv) ||
                            val==std::to_wstring(lv))
                        { ps.headingLevel = lv; break; }
                    }
                }
            } else if ((nm==L"level"||nm==L"headinglevel") && ps.headingLevel > 0) {
                int lv = _wtoi(val.c_str());
                if (lv >= 1 && lv <= 6) ps.headingLevel = lv;
            }
        }
    }
    if (state == IN_CHAR && curId >= 0) chars[curId] = cs;
    if (state == IN_PARA && curId >= 0) paras[curId] = ps;
}

// Parse one HWPX section XML into the CDM builder
void ParseHwpxSection(const std::wstring& xml,
                      const CharShapeMap& charShapes,
                      const ParaShapeMap& paraShapes,
                      cdm::DocumentBuilder& b)
{
    bool paraOpen         = false;
    int  curParaShapeId   = -1;
    int  curCharShapeId   = -1;
    int  paraHeadingLevel = 0;

    // Table state
    bool inTable = false;
    int  inCell  = 0;
    std::wstring cellBuf;
    struct CellEntry { std::wstring text; };
    std::vector<std::vector<CellEntry>> tblRows;
    std::vector<CellEntry>              tblRow;

    auto getPS = [&]() -> const ParaShapeInfo* {
        if (curParaShapeId < 0) return nullptr;
        auto it = paraShapes.find(curParaShapeId);
        return it != paraShapes.end() ? &it->second : nullptr;
    };
    auto getCS = [&]() -> const CharShapeInfo* {
        if (curCharShapeId < 0) return nullptr;
        auto it = charShapes.find(curCharShapeId);
        return it != charShapes.end() ? &it->second : nullptr;
    };

    auto ensurePara = [&]() {
        if (!paraOpen && inCell == 0) {
            b.BeginParagraph();
            if (paraHeadingLevel > 0) {
                b.SetCurrentParagraphStyleRef("Heading" + std::to_string(paraHeadingLevel));
            } else if (const auto* ps = getPS()) {
                if      (ps->alignment == cdm::Alignment::Center)  b.SetCurrentParagraphAlignment(cdm::Alignment::Center);
                else if (ps->alignment == cdm::Alignment::Right)   b.SetCurrentParagraphAlignment(cdm::Alignment::Right);
                else if (ps->alignment == cdm::Alignment::Justify) b.SetCurrentParagraphAlignment(cdm::Alignment::Justify);
            }
            paraOpen = true;
        }
    };

    auto closePara = [&]() {
        if (paraOpen) { b.EndParagraph(); paraOpen = false; }
    };

    auto flushTable = [&]() {
        if (tblRows.empty()) return;
        b.BeginTable();
        for (auto& row : tblRows) {
            b.BeginTableRow();
            for (auto& cell : row) b.AddTableCell(WToU32(cell.text));
            b.EndTableRow();
        }
        b.EndTable();
        tblRows.clear();
    };

    size_t pos = 0;
    while (pos < xml.size()) {
        size_t lt = xml.find(L'<', pos); if (lt == std::wstring::npos) break;
        size_t gt = xml.find(L'>', lt);  if (gt == std::wstring::npos) break;
        const std::wstring raw = xml.substr(lt+1, gt-lt-1);
        pos = gt + 1;
        if (raw.empty() || raw[0]==L'?' || raw[0]==L'!') continue;
        bool isE  = (raw[0] == L'/');
        bool isSC = (!raw.empty() && raw.back() == L'/');
        std::wstring nm = NormTag(raw, isE);

        if (nm==L"p" || nm==L"para") {
            if (!isE && !isSC) {
                auto v = HwpAttr(raw, L"parashapeidref");
                if (v.empty()) v = HwpAttr(raw, L"parastyleidref");
                curParaShapeId   = v.empty() ? -1 : _wtoi(v.c_str());
                paraHeadingLevel = 0;
                if (const auto* ps = getPS()) paraHeadingLevel = ps->headingLevel;
            } else if (isE) {
                if (inCell > 0) { if (!cellBuf.empty() && cellBuf.back()!=L'\n') cellBuf += L'\n'; }
                else closePara();
                curParaShapeId = -1; paraHeadingLevel = 0;
            }
        } else if (nm == L"run") {
            if (!isE && !isSC) {
                auto v = HwpAttr(raw, L"charshapeidref");
                if (v.empty()) v = HwpAttr(raw, L"charstyleidref");
                curCharShapeId = v.empty() ? -1 : _wtoi(v.c_str());
            } else if (isE) { curCharShapeId = -1; }
        } else if (nm == L"t") {
            if (!isE) {
                size_t textEnd = xml.find(L'<', gt+1);
                std::wstring text;
                if (textEnd != std::wstring::npos) { text = xml.substr(gt+1, textEnd-gt-1); pos = textEnd; }
                else { text = xml.substr(gt+1); pos = xml.size(); }
                DecodeXml(text);
                if (text.empty()) continue;

                if (inCell > 0) {
                    cellBuf += text;
                } else {
                    ensurePara();
                    auto u32 = WToU32(text);
                    const CharShapeInfo* cs = getCS();
                    // Don't apply charShape styling to heading paragraphs —
                    // the Heading block rendering in CdmLoader handles sizing.
                    if (cs && cs->HasStyle() && paraHeadingLevel == 0) {
                        cdm::TextStyle st;
                        if (cs->bold)      st.bold      = true;
                        if (cs->italic)    st.italic    = true;
                        if (cs->underline) st.underline = cdm::UnderlineStyle::Single;
                        if (cs->strike)    st.strike    = true;
                        if (cs->color)     st.color     = *cs->color;
                        if (cs->ptSize>0)  st.fontSize  = cdm::Length::Pt(cs->ptSize);
                        b.AddStyledText(u32, st);
                    } else {
                        b.AddText(u32);
                    }
                }
            }
        } else if (nm==L"linebreak" || nm==L"ln" || nm==L"hardreturn") {
            if (inCell > 0) cellBuf += L'\n';
            else if (paraOpen) b.AddLineBreak();
        } else if (nm==L"tbl" || nm==L"table") {
            if (!isE && !isSC) { closePara(); flushTable(); inTable = true; tblRows.clear(); }
            else if (isE) { flushTable(); inTable = false; }
        } else if (nm==L"tr" || nm==L"tablerow") {
            if (!isE && !isSC) tblRow.clear();
            else if (isE) { if (!tblRow.empty()) tblRows.push_back(tblRow); }
        } else if (nm==L"tc" || nm==L"tablecell" || nm==L"cell") {
            if (!isE && !isSC) { inCell++; if (inCell == 1) cellBuf.clear(); }
            else if (isE) {
                inCell--;
                if (inCell == 0) { tblRow.push_back({cellBuf}); cellBuf.clear(); }
            }
        }
    }
    closePara();
    flushTable();
}

} // anonymous namespace

// ── Legacy text-only extractor (kept for fallback) ────────────────────────────

std::wstring HwpxFormat::ParseBodyText(const std::wstring& xml) {
    std::wostringstream out;
    bool firstPara = true;
    size_t pos = 0;
    while (pos < xml.size()) {
        size_t tagStart = xml.find(L'<', pos); if (tagStart == std::wstring::npos) break;
        size_t tagEnd   = xml.find(L'>', tagStart); if (tagEnd == std::wstring::npos) break;
        std::wstring tag = xml.substr(tagStart+1, tagEnd-tagStart-1);
        bool isClose = (!tag.empty() && tag[0] == L'/');
        size_t nameStart = isClose ? 1 : 0;
        size_t sp = tag.find_first_of(L" /\t", nameStart);
        std::wstring rawName = (sp != std::wstring::npos) ? tag.substr(nameStart, sp-nameStart) : tag.substr(nameStart);
        std::wstring name = rawName; for (auto& c : name) c = towlower(c);
        if (!isClose && (name==L"hp:p"||name==L"hh:p")) {
            if (!firstPara) out << L'\n'; firstPara = false;
        } else if (!isClose && (name==L"hp:t"||name==L"hh:t")) {
            std::wstring closeTagL = L"</" + name + L">";
            size_t cs = tagEnd+1;
            size_t ce = FindCI(xml, closeTagL, cs);
            if (ce == std::wstring::npos) { pos = tagEnd+1; continue; }
            std::wstring text = xml.substr(cs, ce-cs);
            DecodeXml(text); out << text;
            pos = ce + closeTagL.size(); continue;
        }
        pos = tagEnd + 1;
    }
    return out.str();
}

// ── Save helpers ──────────────────────────────────────────────────────────────

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
        << "<hsp:Title>"       << XmlEsc(props.title)   << "</hsp:Title>\r\n"
        << "<hsp:Author>"      << XmlEsc(props.author)  << "</hsp:Author>\r\n"
        << "<hsp:Subject>"     << XmlEsc(props.subject) << "</hsp:Subject>\r\n"
        << "<hsp:Description>" << XmlEsc(props.comment) << "</hsp:Description>\r\n"
        << "</hsp:HWPSummaryInfo>\r\n";
    return xml.str();
}

// ── Load ──────────────────────────────────────────────────────────────────────

FormatResult HwpxFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    ZipReader zip;
    if (!zip.Open(path)) { r.error = L"Cannot open HWPX file."; return r; }

    // Metadata
    for (const char* s : {"Contents/summary.xml","Summary.xml","Contents/Summary.xml"}) {
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

    // Parse HWPX header for CharShape/ParaShape definitions
    CharShapeMap charShapes;
    ParaShapeMap paraShapes;
    for (const char* hpath : {
            "Contents/Header.xml",     "Contents/header.xml",
            "Contents/HeaderXML/Header.xml", "Contents/HeaderXML/header.xml",
            "Header.xml" }) {
        if (HasEntryCI(zip, hpath)) {
            ParseHwpxHeader(ExtractWideCI(zip, hpath), charShapes, paraShapes);
            break;
        }
    }

    // Parse section files into CDM
    cdm::DocumentBuilder b;
    b.SetOriginalFormat(cdm::FileFormat::HWPX);
    b.BeginSection();

    bool gotContent = false;
    for (int i = 0; i < 100; ++i) {
        std::string found;
        for (const char* fmt : {
                "Contents/Sections/Section%d.xml",
                "Contents/sections/section%d.xml",
                "BodyText/Section%d.xml",
                "bodytext/section%d.xml",
                "Contents/section%d.xml" }) {
            char buf[64]; snprintf(buf, sizeof(buf), fmt, i);
            std::string real = FindEntryCI(zip, buf);
            if (!real.empty()) { found = real; break; }
        }
        if (found.empty()) break;
        ParseHwpxSection(zip.ExtractWide(found), charShapes, paraShapes, b);
        gotContent = true;
    }

    b.EndSection();
    if (!gotContent) { r.error = L"HWPX text extraction failed."; return r; }
    r.cdmDoc = b.MoveBuild();
    r.ok = true;
    return r;
}

// ── Save ──────────────────────────────────────────────────────────────────────

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
