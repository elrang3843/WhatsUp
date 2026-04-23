#include "HwpFormat.h"
#include "../i18n/Localization.h"
#include <objbase.h>
#include <objidl.h>
#include <sstream>
#include <vector>
#include <cstdint>
#include "zlib.h"

// HWP 5.x uses OLE2 Compound Document Format.
// BodyText/SectionN streams are zlib-compressed (unless FileHeader flags say otherwise).
// Each record: 4-byte header [TagID:10 | Level:4 | Size:18]
// TagID 66 = paragraph, TagID 67 = character (UTF-16LE wchar_t).

static std::vector<uint8_t> ReadFileBytes(const std::wstring& path) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
                           nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return {};
    DWORD sz = GetFileSize(h, nullptr);
    std::vector<uint8_t> buf(sz);
    DWORD rd = 0;
    ReadFile(h, buf.data(), sz, &rd, nullptr);
    CloseHandle(h);
    buf.resize(rd);
    return buf;
}

// zlib inflate: handles both zlib (0x78xx header) and raw deflate
static std::vector<uint8_t> InflateData(const std::vector<uint8_t>& src) {
    if (src.empty()) return {};
    std::vector<uint8_t> out;
    out.reserve(src.size() * 4);

    z_stream z{};
    // MAX_WBITS+32: auto-detect zlib or gzip header;
    // on failure fall back to raw deflate (negative wbits)
    int ret = inflateInit2(&z, MAX_WBITS + 32);
    if (ret != Z_OK) {
        ret = inflateInit2(&z, -MAX_WBITS);
        if (ret != Z_OK) return {};
    }

    z.avail_in  = static_cast<uInt>(src.size());
    z.next_in   = const_cast<Bytef*>(src.data());

    do {
        size_t off = out.size();
        out.resize(off + 65536);
        z.avail_out = 65536;
        z.next_out  = out.data() + off;
        ret = inflate(&z, Z_NO_FLUSH);
        out.resize(off + 65536 - z.avail_out);
    } while (ret == Z_OK && z.avail_in > 0);

    inflateEnd(&z);
    return (ret == Z_STREAM_END || ret == Z_OK) ? out : std::vector<uint8_t>{};
}

std::wstring HwpFormat::ExtractText(const std::vector<uint8_t>& data) {
    std::wostringstream out;
    size_t i = 0;
    bool firstPara = true;

    while (i + 4 <= data.size()) {
        uint32_t hdr = static_cast<uint32_t>(data[i])
                     | (static_cast<uint32_t>(data[i+1]) << 8)
                     | (static_cast<uint32_t>(data[i+2]) << 16)
                     | (static_cast<uint32_t>(data[i+3]) << 24);
        i += 4;

        uint32_t tagId = hdr & 0x3FF;
        uint32_t recSz = (hdr >> 20) & 0xFFF;
        if (recSz == 0xFFF) {
            if (i + 4 > data.size()) break;
            recSz = static_cast<uint32_t>(data[i])
                  | (static_cast<uint32_t>(data[i+1]) << 8)
                  | (static_cast<uint32_t>(data[i+2]) << 16)
                  | (static_cast<uint32_t>(data[i+3]) << 24);
            i += 4;
        }
        if (i + recSz > data.size()) break;

        if (tagId == 66) {  // HWPTAG_PARA_HEADER
            if (!firstPara) out << L'\n';
            firstPara = false;
        } else if (tagId == 67) {  // HWPTAG_CHAR (UTF-16LE)
            for (size_t j = 0; j + 1 < recSz; j += 2) {
                wchar_t ch = static_cast<wchar_t>(data[i + j])
                           | (static_cast<wchar_t>(data[i + j + 1]) << 8);
                if (ch >= 0x20 || ch == L'\t')
                    out << ch;
            }
        }
        i += recSz;
    }
    return out.str();
}

std::vector<uint8_t> HwpFormat::ReadOleStream(const std::wstring& path,
                                               const std::wstring& streamName) {
    IStorage* pStorage = nullptr;
    HRESULT hr = StgOpenStorage(path.c_str(), nullptr,
                                STGM_READ | STGM_SHARE_DENY_WRITE,
                                nullptr, 0, &pStorage);
    if (FAILED(hr) || !pStorage) return {};

    IStorage* pBodyText = nullptr;
    hr = pStorage->OpenStorage(L"BodyText", nullptr,
                               STGM_READ | STGM_SHARE_EXCLUSIVE,
                               nullptr, 0, &pBodyText);

    IStream* pStream = nullptr;
    if (SUCCEEDED(hr) && pBodyText) {
        hr = pBodyText->OpenStream(streamName.c_str(), nullptr,
                                   STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    } else {
        hr = pStorage->OpenStream(streamName.c_str(), nullptr,
                                  STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    }

    std::vector<uint8_t> data;
    if (SUCCEEDED(hr) && pStream) {
        STATSTG stat{};
        pStream->Stat(&stat, STATFLAG_NONAME);
        ULONG sz = static_cast<ULONG>(stat.cbSize.QuadPart);
        data.resize(sz);
        ULONG read = 0;
        pStream->Read(data.data(), sz, &read);
        data.resize(read);
        pStream->Release();
    }

    if (pBodyText) pBodyText->Release();
    pStorage->Release();
    return data;
}

// Read FileHeader stream to determine if BodyText is compressed
static bool IsHwpBodyCompressed(const std::wstring& path) {
    IStorage* pStorage = nullptr;
    if (FAILED(StgOpenStorage(path.c_str(), nullptr,
                              STGM_READ | STGM_SHARE_DENY_WRITE,
                              nullptr, 0, &pStorage)) || !pStorage)
        return true; // assume compressed if we can't check

    IStream* pStream = nullptr;
    pStorage->OpenStream(L"FileHeader", nullptr,
                         STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);

    bool compressed = true;
    if (pStream) {
        // FileHeader: 32-byte signature + 4-byte version + 4-byte flags
        uint8_t hdr[40] = {};
        ULONG read = 0;
        pStream->Read(hdr, sizeof(hdr), &read);
        if (read >= 36) {
            uint32_t flags = static_cast<uint32_t>(hdr[32])
                           | (static_cast<uint32_t>(hdr[33]) << 8)
                           | (static_cast<uint32_t>(hdr[34]) << 16)
                           | (static_cast<uint32_t>(hdr[35]) << 24);
            compressed = (flags & 1) != 0;
        }
        pStream->Release();
    }
    pStorage->Release();
    return compressed;
}

FormatResult HwpFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    doc.SetReadOnly(true);

    auto header = ReadFileBytes(path);
    if (header.size() < 4) { r.error = L"Invalid HWP file."; return r; }

    bool isOle = (header[0] == 0xD0 && header[1] == 0xCF
               && header[2] == 0x11 && header[3] == 0xE0);
    if (!isOle) {
        r.error = L"HWP format not supported (only HWP 5.x OLE files).";
        return r;
    }

    bool compressed = IsHwpBodyCompressed(path);
    std::wostringstream text;

    for (int sec = 0; sec < 100; ++sec) {
        std::wstring secName = L"Section" + std::to_wstring(sec);
        auto stream = ReadOleStream(path, secName);
        if (stream.empty()) break;

        std::vector<uint8_t> data;
        if (compressed) {
            data = InflateData(stream);
            if (data.empty()) data = stream; // fallback: try raw
        } else {
            data = stream;
        }

        std::wstring secText = ExtractText(data);
        if (!secText.empty()) text << secText << L'\n';
    }

    std::wstring result = text.str();
    if (result.empty()) {
        r.error = L"HWP text extraction failed (only HWP 5.x supported, read-only).";
        return r;
    }

    r.content = result;
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult HwpFormat::Save(const std::wstring& /*path*/,
                              const std::wstring& /*content*/,
                              const std::string&  /*rtf*/,
                              Document& /*doc*/) {
    FormatResult r;
    r.error = Localization::Get(StrID::MSG_HWP_READONLY);
    return r;
}
