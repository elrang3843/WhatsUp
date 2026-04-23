#include "Document.h"
#include "i18n/Localization.h"
#include <shlwapi.h>
#include <algorithm>
#include <cctype>

Document::Document() {
    Reset();
}

void Document::Reset() {
    m_path.clear();
    m_format   = DocFormat::Txt;
    m_modified = false;
    m_readOnly = false;
    m_props    = {};
    m_stats    = {};
}

std::wstring Document::DisplayTitle() const {
    std::wstring name;
    if (m_path.empty()) {
        name = Localization::Get(SID::UNTITLED);
    } else {
        wchar_t fname[MAX_PATH];
        wcscpy_s(fname, m_path.c_str());
        name = PathFindFileNameW(fname);
    }
    if (m_modified) name = L"*" + name;
    return name;
}

DocFormat Document::DetectFormat(const std::wstring& path) {
    if (path.empty()) return DocFormat::Txt;

    const wchar_t* ext = PathFindExtensionW(path.c_str());
    if (!ext || !*ext) return DocFormat::Txt;

    // Case-insensitive comparison
    std::wstring e = ext;
    std::transform(e.begin(), e.end(), e.begin(), ::towlower);

    if (e == L".txt")              return DocFormat::Txt;
    if (e == L".html" || e == L".htm") return DocFormat::Html;
    if (e == L".md")               return DocFormat::Md;
    if (e == L".docx")             return DocFormat::Docx;
    if (e == L".hwpx")             return DocFormat::Hwpx;
    if (e == L".hwp")              return DocFormat::Hwp;
    if (e == L".doc")              return DocFormat::Doc;
    return DocFormat::Txt;
}

const wchar_t* Document::FormatName(DocFormat f) {
    switch (f) {
        case DocFormat::Txt:  return L"TXT";
        case DocFormat::Html: return L"HTML";
        case DocFormat::Md:   return L"Markdown";
        case DocFormat::Docx: return L"DOCX";
        case DocFormat::Hwpx: return L"HWPX";
        case DocFormat::Hwp:  return L"HWP";
        case DocFormat::Doc:  return L"DOC";
        default:              return L"Unknown";
    }
}
