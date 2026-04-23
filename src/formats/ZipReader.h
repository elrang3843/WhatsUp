#pragma once
// Minimal ZIP reader for DOCX/HWPX support.
// Only handles stored (0) and deflate (8) compression methods.
// Requires zlib (fetched via CMake FetchContent).
#include <windows.h>
#include <string>
#include <vector>
#include <unordered_map>
#include <cstdint>
#include "zlib.h"

struct ZipEntry {
    std::string  name;
    uint32_t     method;         // 0 = stored, 8 = deflate
    uint32_t     crc32;
    uint32_t     compressedSize;
    uint32_t     uncompressedSize;
    uint64_t     offset;         // offset of local file header
};

class ZipReader {
public:
    bool        Open(const std::wstring& path);
    void        Close();

    bool        HasEntry(const std::string& name) const;
    // Extract entry to a byte buffer; returns empty on failure
    std::vector<uint8_t> Extract(const std::string& name);
    // Extract entry as UTF-8 string
    std::string ExtractText(const std::string& name);
    // Extract entry as UTF-16 string (converts from UTF-8)
    std::wstring ExtractWide(const std::string& name);

    const std::unordered_map<std::string, ZipEntry>& Entries() const { return m_entries; }

    ~ZipReader() { Close(); }

private:
    bool ParseCentralDirectory();

    HANDLE                                       m_hFile = INVALID_HANDLE_VALUE;
    std::unordered_map<std::string, ZipEntry>    m_entries;
};

// ---- ZipWriter: creates a ZIP file ----
class ZipWriter {
public:
    bool  Open(const std::wstring& path);
    bool  AddFile(const std::string& entryName, const std::vector<uint8_t>& data,
                  bool deflate = true);
    bool  AddText(const std::string& entryName, const std::string& utf8Text,
                  bool deflate = true);
    bool  Close();
    ~ZipWriter() { Close(); }

private:
    struct LocalRecord {
        std::string name;
        uint32_t    crc32;
        uint32_t    compressedSize;
        uint32_t    uncompressedSize;
        uint32_t    method;
        uint32_t    localHeaderOffset;
    };

    HANDLE               m_hFile = INVALID_HANDLE_VALUE;
    std::vector<LocalRecord> m_records;

    bool WriteLocalHeader(const LocalRecord& rec, const std::vector<uint8_t>& compressed);
    static uint32_t Crc32(const uint8_t* data, size_t size);
};
