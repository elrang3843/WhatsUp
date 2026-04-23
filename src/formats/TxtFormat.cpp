#include "TxtFormat.h"
#include <fstream>
#include <sstream>
#include <vector>
#include <stdexcept>

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
        // UTF-16 LE BOM
        if (bytes[0] == 0xFF && bytes[1] == 0xFE) {
            std::wstring ws(reinterpret_cast<const wchar_t*>(bytes.data() + 2),
                            (bytes.size() - 2) / 2);
            return ws;
        }
        // UTF-16 BE BOM — convert to LE
        if (bytes[0] == 0xFE && bytes[1] == 0xFF) {
            std::wstring ws;
            for (size_t i = 2; i + 1 < bytes.size(); i += 2) {
                wchar_t ch = (static_cast<wchar_t>(bytes[i]) << 8) | bytes[i + 1];
                ws += ch;
            }
            return ws;
        }
    }
    // UTF-8 (with or without BOM)
    size_t start = 0;
    if (bytes.size() >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        start = 3;

    int needed = MultiByteToWideChar(CP_UTF8, 0,
                                     reinterpret_cast<LPCCH>(bytes.data() + start),
                                     static_cast<int>(bytes.size() - start),
                                     nullptr, 0);
    if (needed <= 0) {
        // Fallback: try system ANSI code page
        needed = MultiByteToWideChar(CP_ACP, 0,
                                     reinterpret_cast<LPCCH>(bytes.data()),
                                     static_cast<int>(bytes.size()),
                                     nullptr, 0);
        std::wstring ws(needed, 0);
        MultiByteToWideChar(CP_ACP, 0,
                            reinterpret_cast<LPCCH>(bytes.data()),
                            static_cast<int>(bytes.size()),
                            ws.data(), needed);
        return ws;
    }
    std::wstring ws(needed, 0);
    MultiByteToWideChar(CP_UTF8, 0,
                        reinterpret_cast<LPCCH>(bytes.data() + start),
                        static_cast<int>(bytes.size() - start),
                        ws.data(), needed);
    return ws;
}

FormatResult TxtFormat::Load(const std::wstring& path, Document& /*doc*/) {
    FormatResult r;
    auto bytes = ReadFile(path);
    if (bytes.empty() && GetLastError() != ERROR_SUCCESS) {
        r.error = L"Cannot read file.";
        return r;
    }
    r.content = DecodeText(bytes);
    r.rtf     = false;
    r.ok      = true;
    return r;
}

FormatResult TxtFormat::Save(const std::wstring& path,
                              const std::wstring& content,
                              const std::string&  /*rtf*/,
                              Document& /*doc*/) {
    FormatResult r;

    // Write UTF-8 with BOM
    HANDLE hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0,
                               nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        r.error = L"Cannot create file.";
        return r;
    }

    // UTF-8 BOM
    const uint8_t bom[3] = { 0xEF, 0xBB, 0xBF };
    DWORD written = 0;
    WriteFile(hFile, bom, 3, &written, nullptr);

    int needed = WideCharToMultiByte(CP_UTF8, 0, content.c_str(),
                                     static_cast<int>(content.size()),
                                     nullptr, 0, nullptr, nullptr);
    if (needed > 0) {
        std::vector<char> utf8(needed);
        WideCharToMultiByte(CP_UTF8, 0, content.c_str(),
                            static_cast<int>(content.size()),
                            utf8.data(), needed, nullptr, nullptr);
        WriteFile(hFile, utf8.data(), needed, &written, nullptr);
    }
    CloseHandle(hFile);
    r.ok = true;
    return r;
}
