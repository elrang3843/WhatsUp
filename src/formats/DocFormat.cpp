#include "DocFormat.h"
#include "../i18n/Localization.h"
#include <objbase.h>
#include <objidl.h>
#include <sstream>
#include <vector>
#include <cstdint>

static const std::string kRtfHeader =
    "{\\rtf1\\ansi\\deff2\n"
    "{\\fonttbl\n"
    "{\\f0\\froman\\fcharset0 Times New Roman;}\n"
    "{\\f1\\fswiss\\fcharset0 Calibri;}\n"
    "{\\f2\\fswiss\\fcharset129 Malgun Gothic;}\n"
    "{\\f3\\fmodern\\fcharset0 Courier New;}\n"
    "}\n"
    "\\f2\\fs22\\pard\\ql\n";

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

static std::string TextToRtf(const std::wstring& text) {
    std::string body;
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        body += "\\pard\\ql ";
        body += RtfEnc(line);
        body += "\\par\n";
    }
    return kRtfHeader + body + "}";
}

// DOC format: OLE2 Compound Document (Word 97-2003).
// Text is located via the Clx (piece table) stored in the 1Table/0Table stream.
// FIB field fcClx is at offset 418 in the WordDocument stream (FibRgFcLcb97[33]).
// Each piece descriptor (PCD) has an FcCompressed field: bit 30 = isAnsi,
// lower 30 bits = byte offset in WordDocument (divide by 2 when isAnsi).

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
    auto wd = ReadOleStream(path, L"WordDocument");
    if (wd.size() < 426) return {};

    uint16_t magic = LE16(wd.data());
    if (magic != 0xA5EC && magic != 0xA5DC) return {};

    // FIB offset 10-11: flags; bit 9 (0x0200) = fWhichTblStm (0=0Table, 1=1Table)
    uint16_t flags = LE16(wd.data() + 10);
    bool useTable1 = (flags & 0x0200) != 0;

    // Clx location: FibRgFcLcb97 starts at 154; entry [33] = Clx → offset 154 + 33*8 = 418
    uint32_t fcClx  = LE32(wd.data() + 418);
    uint32_t lcbClx = LE32(wd.data() + 422);
    if (lcbClx == 0) return {};

    auto tbl = ReadOleStream(path, useTable1 ? L"1Table" : L"0Table");
    if (fcClx + lcbClx > tbl.size()) return {};

    const uint8_t* clx    = tbl.data() + fcClx;
    size_t         clxPos = 0;

    // Skip grpprl entries (type byte 0x01)
    while (clxPos < lcbClx && clx[clxPos] == 0x01) {
        clxPos++;
        if (clxPos + 2 > lcbClx) return {};
        uint16_t cb = LE16(clx + clxPos);
        clxPos += 2 + cb;
    }

    // Piece table entry must start with type byte 0x02
    if (clxPos >= lcbClx || clx[clxPos] != 0x02) return {};
    clxPos++;
    if (clxPos + 4 > lcbClx) return {};
    uint32_t lcbPcd = LE32(clx + clxPos);
    clxPos += 4;
    if (lcbPcd < 4 || clxPos + lcbPcd > lcbClx) return {};

    // PlcfPcd: (n+1) CP values (4 bytes each) + n PCD values (8 bytes each)
    // 4*(n+1) + 8*n = lcbPcd  →  n = (lcbPcd - 4) / 12
    uint32_t nPieces = (lcbPcd - 4) / 12;
    if (nPieces == 0) return {};

    const uint8_t* pCPs  = clx + clxPos;
    const uint8_t* pPCDs = pCPs + (nPieces + 1) * 4;

    std::wostringstream out;
    for (uint32_t i = 0; i < nPieces; i++) {
        uint32_t cpStart = LE32(pCPs + i * 4);
        uint32_t cpEnd   = LE32(pCPs + (i + 1) * 4);
        if (cpEnd <= cpStart) continue;
        uint32_t nCh = cpEnd - cpStart;

        // PCD: 2-byte flags | 4-byte FcCompressed | 2-byte prm
        const uint8_t* pcd = pPCDs + i * 8;
        uint32_t fcRaw  = LE32(pcd + 2);
        bool isAnsi     = (fcRaw & 0x40000000) != 0; // bit 30
        uint32_t fc     = fcRaw & 0x3FFFFFFF;
        uint32_t off    = isAnsi ? fc / 2 : fc;

        if (isAnsi) {
            if (off + nCh > wd.size()) continue;
            for (uint32_t j = 0; j < nCh; j++) {
                uint8_t c = wd[off + j];
                if      (c == 0x0D)              out << L'\n';
                else if (c == 0x07 || c == 0x09) out << L'\t';
                else if (c >= 0x20)              out << static_cast<wchar_t>(c);
            }
        } else {
            if (off + (uint64_t)nCh * 2 > wd.size()) continue;
            for (uint32_t j = 0; j < nCh; j++) {
                wchar_t ch = static_cast<wchar_t>(wd[off + j * 2])
                           | (static_cast<wchar_t>(wd[off + j * 2 + 1]) << 8);
                if      (ch == 0x0D)                    out << L'\n';
                else if (ch == 0x07 || ch == 0x09)      out << L'\t';
                else if (ch >= 0x20 && ch < 0xFFFE)     out << ch;
            }
        }
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
    std::string rtf = TextToRtf(text);
    r.content = std::wstring(rtf.begin(), rtf.end());
    r.rtf     = true;
    r.ok      = true;
    return r;
}

FormatResult DocFormat_::Save(const std::wstring& /*path*/,
                               const std::wstring& /*content*/,
                               const std::string&  /*rtf*/,
                               Document& /*doc*/) {
    FormatResult r;
    r.error = Localization::Get(StrID::MSG_DOC_READONLY);
    return r;
}
