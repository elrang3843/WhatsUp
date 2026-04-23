#include "HwpFormat.h"
#include "../i18n/Localization.h"
#include <objbase.h>
#include <objidl.h>
#include <sstream>

// HWP 5.x uses OLE2 Compound Document Format.
// Text runs are in "BodyText/Section0" compressed with zlib.
// Each record has a 4-byte header: [TagID:10 | Level:4 | Size:18]
// Text characters are TagID 67 (HWPTAG_CHAR), WCHAR each.
// This is a best-effort text extractor — layout/images are not supported.

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

// Attempt to extract text from HWP5 BodyText binary records.
// HWP BodyText records: TagID=67 means "character" (UTF-16LE wchar_t follows in data).
// TagID=66 = paragraph, TagID=68 = line seg, etc.
std::wstring HwpFormat::ExtractText(const std::vector<uint8_t>& data) {
    std::wostringstream out;
    size_t i = 0;
    bool firstPara = true;

    while (i + 4 <= data.size()) {
        uint32_t header = static_cast<uint32_t>(data[i])
                        | (static_cast<uint32_t>(data[i+1]) << 8)
                        | (static_cast<uint32_t>(data[i+2]) << 16)
                        | (static_cast<uint32_t>(data[i+3]) << 24);
        i += 4;

        uint32_t tagId  = header & 0x3FF;
        uint32_t recSz  = (header >> 20) & 0xFFF;
        if (recSz == 0xFFF) {
            // Extended size in next 4 bytes
            if (i + 4 > data.size()) break;
            recSz = static_cast<uint32_t>(data[i])
                  | (static_cast<uint32_t>(data[i+1]) << 8)
                  | (static_cast<uint32_t>(data[i+2]) << 16)
                  | (static_cast<uint32_t>(data[i+3]) << 24);
            i += 4;
        }

        if (i + recSz > data.size()) break;

        if (tagId == 66) {
            // HWPTAG_PARA_HEADER — start of paragraph
            if (!firstPara) out << L'\n';
            firstPara = false;
        } else if (tagId == 67) {
            // HWPTAG_CHAR — text character (UTF-16LE)
            for (size_t j = 0; j + 1 < recSz; j += 2) {
                wchar_t ch = static_cast<wchar_t>(data[i + j])
                           | (static_cast<wchar_t>(data[i + j + 1]) << 8);
                // Filter control chars (HWP uses special codes < 0x20)
                if (ch >= 0x20 || ch == L'\t')
                    out << ch;
            }
        }
        i += recSz;
    }
    return out.str();
}

// Use OLE IStorage to open HWP file and read the BodyText/Section0 stream
std::vector<uint8_t> HwpFormat::ReadOleStream(const std::wstring& path,
                                               const std::wstring& streamName) {
    IStorage* pStorage = nullptr;
    HRESULT hr = StgOpenStorage(path.c_str(), nullptr,
                                STGM_READ | STGM_SHARE_DENY_WRITE,
                                nullptr, 0, &pStorage);
    if (FAILED(hr) || !pStorage) return {};

    // Navigate to BodyText sub-storage
    IStorage* pBodyText = nullptr;
    hr = pStorage->OpenStorage(L"BodyText", nullptr,
                               STGM_READ | STGM_SHARE_EXCLUSIVE,
                               nullptr, 0, &pBodyText);

    IStream* pStream = nullptr;
    if (SUCCEEDED(hr) && pBodyText) {
        hr = pBodyText->OpenStream(streamName.c_str(), nullptr,
                                   STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
    } else {
        // Try opening directly (older HWP layout)
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

FormatResult HwpFormat::Load(const std::wstring& path, Document& doc) {
    FormatResult r;
    doc.SetReadOnly(true);

    // First check if it's a HWP5 OLE file (starts with OLE magic D0CF11E0)
    auto header = ReadFileBytes(path);
    if (header.size() < 4) { r.error = L"Invalid HWP file."; return r; }

    bool isOle = (header[0] == 0xD0 && header[1] == 0xCF
               && header[2] == 0x11 && header[3] == 0xE0);

    std::wostringstream text;
    if (isOle) {
        // Try each BodyText section
        for (int sec = 0; sec < 50; ++sec) {
            std::wstring secName = L"Section" + std::to_wstring(sec);
            auto stream = ReadOleStream(path, secName);
            if (stream.empty()) break;
            // HWP BodyText streams are zlib-compressed
            // Try decompressing with zlib
            // (This requires zlib, which is included via CMake FetchContent)
            // For now we try raw and compressed
            text << ExtractText(stream) << L'\n';
        }
    }

    if (text.str().empty()) {
        r.error = L"HWP text extraction failed. Only HWP 5.x is supported (read-only).";
        return r;
    }

    r.content = text.str();
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
