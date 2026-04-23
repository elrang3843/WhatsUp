#include "DocFormat.h"
#include "../i18n/Localization.h"
#include <objbase.h>
#include <objidl.h>
#include <sstream>

// DOC format uses OLE2 Compound Document.
// The main text is in the "WordDocument" stream.
// The File Information Block (FIB) at offset 0 contains fcMin and fcMac (text range).
// Text content starts at fcMin as UTF-16LE characters (for Word97+ Unicode docs).
// This is a simplified extractor — only handles Unicode text content.

static std::vector<uint8_t> ReadOleStream(const std::wstring& path,
                                          const wchar_t* streamName) {
    IStorage* pStorage = nullptr;
    if (FAILED(StgOpenStorage(path.c_str(), nullptr,
                              STGM_READ | STGM_SHARE_DENY_WRITE,
                              nullptr, 0, &pStorage)) || !pStorage)
        return {};

    IStream* pStream = nullptr;
    pStorage->OpenStream(streamName, nullptr,
                         STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);

    std::vector<uint8_t> data;
    if (pStream) {
        STATSTG stat{};
        pStream->Stat(&stat, STATFLAG_NONAME);
        ULONG sz = static_cast<ULONG>(stat.cbSize.QuadPart);
        data.resize(sz);
        ULONG read = 0;
        pStream->Read(data.data(), sz, &read);
        data.resize(read);
        pStream->Release();
    }
    pStorage->Release();
    return data;
}

static uint32_t LE32(const uint8_t* p) {
    return static_cast<uint32_t>(p[0])
         | (static_cast<uint32_t>(p[1]) << 8)
         | (static_cast<uint32_t>(p[2]) << 16)
         | (static_cast<uint32_t>(p[3]) << 24);
}
static uint16_t LE16(const uint8_t* p) {
    return static_cast<uint16_t>(p[0]) | (static_cast<uint16_t>(p[1]) << 8);
}

std::wstring DocFormat_::ExtractText(const std::wstring& path) {
    // Read WordDocument stream
    auto wdStream = ReadOleStream(path, L"WordDocument");
    if (wdStream.size() < 0x20) return {};

    // FIB: check magic (Word 97+ = 0xA5EC or 0xA5DC)
    uint16_t magic = LE16(wdStream.data());
    if (magic != 0xA5EC && magic != 0xA5DC) return {};

    // FIB fields (offsets per MS-DOC spec)
    // fcMin: byte offset of main text in 1Table/0Table stream (0=ANSI, 1=Unicode)
    // fib.base: 32 bytes, fib.fibRgW97 starts at 32
    // For simplified extraction: read all WCHAR from the 1Table "main text" piece
    // Actually, text content for Word97+ unicode is directly in the document text range.

    // Read 1Table stream for piece table (Clx)
    bool useUnicode = (wdStream.size() > 10) && ((wdStream[10] & 0x80) != 0);
    auto tblStream = ReadOleStream(path, useUnicode ? L"1Table" : L"0Table");

    // A very simplified approach: scan WordDocument for plausible Unicode text
    // (handles most simple documents correctly)
    std::wostringstream out;
    const uint8_t* d = wdStream.data();
    size_t sz = wdStream.size();

    // Try to find text using FIB: fcPlcfBtePapx gives paragraph positions
    // For simplicity, scan the stream for printable Unicode runs
    for (size_t i = 512; i + 1 < sz; i += 2) {
        wchar_t ch = static_cast<wchar_t>(d[i]) | (static_cast<wchar_t>(d[i + 1]) << 8);
        if (ch == 0x0D) { out << L'\n'; continue; }
        if (ch == 0x07) { out << L'\t'; continue; } // table cell
        if (ch >= 0x20 && ch < 0xFFFE) out << ch;
        else if (ch == 0x09) out << L'\t';
    }
    return out.str();
}

FormatResult DocFormat_::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    doc.SetReadOnly(true);

    std::wstring text = ExtractText(path);
    if (text.empty()) {
        r.error = L"DOC text extraction failed. Only Word 97-2003 Unicode documents are supported (read-only).";
        return r;
    }
    r.content = text;
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult DocFormat_::Save(const std::wstring& /*path*/,
                               const std::wstring& /*content*/,
                               const std::string&  /*rtf*/,
                               Document& /*doc*/) {
    FormatResult r;
    r.error = Localization::Get(SID::MSG_DOC_READONLY);
    return r;
}
