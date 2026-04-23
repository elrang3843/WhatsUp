#include "TxtFormat.h"
#include "../cdm/document_builder.hpp"
#include <sstream>
#include <vector>

// Read bytes from file
static std::vector<uint8_t> ReadFile(const std::wstring& path) {
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
                               nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) return {};
    DWORD size = GetFileSize(hFile, nullptr);
    std::vector<uint8_t> buf(size);
    DWORD read = 0;
    ReadFile(hFile, buf.data(), size, &read, nullptr);
    CloseHandle(hFile);
    buf.resize(read);
    return buf;
}

// Detect BOM and decode to UTF-16
static std::wstring DecodeText(const std::vector<uint8_t>& bytes) {
    if (bytes.size() >= 2) {
        if (bytes[0] == 0xFF && bytes[1] == 0xFE) {
            return std::wstring(reinterpret_cast<const wchar_t*>(bytes.data() + 2),
                                (bytes.size() - 2) / 2);
        }
        if (bytes[0] == 0xFE && bytes[1] == 0xFF) {
            std::wstring ws;
            for (size_t i = 2; i + 1 < bytes.size(); i += 2)
                ws += static_cast<wchar_t>((bytes[i] << 8) | bytes[i+1]);
            return ws;
        }
    }
    size_t start = 0;
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        start = 3;

    int needed = MultiByteToWideChar(CP_UTF8, 0,
        reinterpret_cast<LPCCH>(bytes.data() + start),
        static_cast<int>(bytes.size() - start), nullptr, 0);
    if (needed <= 0) {
        needed = MultiByteToWideChar(CP_ACP, 0,
            reinterpret_cast<LPCCH>(bytes.data()),
            static_cast<int>(bytes.size()), nullptr, 0);
        std::wstring ws(needed, 0);
        MultiByteToWideChar(CP_ACP, 0,
            reinterpret_cast<LPCCH>(bytes.data()),
            static_cast<int>(bytes.size()), ws.data(), needed);
        return ws;
    }
    std::wstring ws(needed, 0);
    MultiByteToWideChar(CP_UTF8, 0,
        reinterpret_cast<LPCCH>(bytes.data() + start),
        static_cast<int>(bytes.size() - start), ws.data(), needed);
    return ws;
}

static std::u32string ToU32(const std::wstring& ws) {
    std::u32string s;
    s.reserve(ws.size());
    for (wchar_t c : ws) s += static_cast<char32_t>(static_cast<uint16_t>(c));
    return s;
}

static cdm::Document TextToCdm(const std::wstring& text, cdm::FileFormat fmt) {
    cdm::DocumentBuilder b;
    b.SetOriginalFormat(fmt);
    b.BeginSection();
    std::wistringstream in(text);
    std::wstring line;
    while (std::getline(in, line)) {
        if (!line.empty() && line.back() == L'\r') line.pop_back();
        b.AddParagraph(ToU32(line));
    }
    b.EndSection();
    return b.MoveBuild();
}

FormatResult TxtFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    auto bytes = ReadFile(path);
    if (bytes.empty() && GetLastError() != ERROR_SUCCESS) {
        r.error = L"Cannot read file.";
        return r;
    }
    std::wstring text = DecodeText(bytes);
    r.cdmDoc = TextToCdm(text, cdm::FileFormat::TXT);
    r.ok     = true;
    return r;
}

FormatResult TxtFormat::Save(const std::wstring& path,
                              const std::wstring& content,
                              const std::string&  /*rtf*/,
                              Document& /*doc*/) {
    FormatResult r;

    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0,
                               nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        r.error = L"Cannot create file.";
        return r;
    }

    const uint8_t bom[3] = { 0xEF, 0xBB, 0xBF };
    DWORD written = 0;
    WriteFile(hFile, bom, 3, &written, nullptr);

    int needed = WideCharToMultiByte(CP_UTF8, 0, content.c_str(),
        static_cast<int>(content.size()), nullptr, 0, nullptr, nullptr);
    if (needed > 0) {
        std::vector<char> utf8(needed);
        WideCharToMultiByte(CP_UTF8, 0, content.c_str(),
            static_cast<int>(content.size()), utf8.data(), needed, nullptr, nullptr);
        WriteFile(hFile, utf8.data(), needed, &written, nullptr);
    }
    CloseHandle(hFile);
    r.ok = true;
    return r;
}
