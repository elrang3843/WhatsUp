#include "ZipReader.h"
#include <cstring>
#include <algorithm>

// --- Low-level helpers ---

static bool ReadAt(HANDLE h, uint64_t offset, void* buf, DWORD size) {
    LARGE_INTEGER li;
    li.QuadPart = static_cast<LONGLONG>(offset);
    if (!SetFilePointerEx(h, li, nullptr, FILE_BEGIN)) return false;
    DWORD read = 0;
    return ReadFile(h, buf, size, &read, nullptr) && read == size;
}

static uint16_t LE16(const uint8_t* p) {
    return static_cast<uint16_t>(p[0]) | (static_cast<uint16_t>(p[1]) << 8);
}
static uint32_t LE32(const uint8_t* p) {
    return static_cast<uint32_t>(p[0])
         | (static_cast<uint32_t>(p[1]) <<  8)
         | (static_cast<uint32_t>(p[2]) << 16)
         | (static_cast<uint32_t>(p[3]) << 24);
}

// --- ZipReader ---

bool ZipReader::Open(const std::wstring& path) {
    Close();
    m_hFile = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ,
                          nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (m_hFile == INVALID_HANDLE_VALUE) return false;
    return ParseCentralDirectory();
}

void ZipReader::Close() {
    if (m_hFile != INVALID_HANDLE_VALUE) {
        CloseHandle(m_hFile);
        m_hFile = INVALID_HANDLE_VALUE;
    }
    m_entries.clear();
}

bool ZipReader::ParseCentralDirectory() {
    // Find End of Central Directory record
    LARGE_INTEGER fileSize{};
    if (!GetFileSizeEx(m_hFile, &fileSize)) return false;
    int64_t size = fileSize.QuadPart;

    // EOCD is at least 22 bytes from end; search backwards for signature 0x06054b50
    const int64_t maxSearch = std::min<int64_t>(size, 65536 + 22);
    std::vector<uint8_t> tail(static_cast<size_t>(maxSearch));
    int64_t tailStart = size - maxSearch;
    if (!ReadAt(m_hFile, static_cast<uint64_t>(tailStart), tail.data(), static_cast<DWORD>(maxSearch)))
        return false;

    int64_t eocdOffset = -1;
    for (int64_t i = maxSearch - 22; i >= 0; --i) {
        if (tail[i] == 0x50 && tail[i+1] == 0x4B && tail[i+2] == 0x05 && tail[i+3] == 0x06) {
            eocdOffset = tailStart + i;
            break;
        }
    }
    if (eocdOffset < 0) return false;

    uint8_t eocd[22];
    if (!ReadAt(m_hFile, static_cast<uint64_t>(eocdOffset), eocd, 22)) return false;

    uint32_t cdOffset = LE32(eocd + 16);
    uint32_t cdSize   = LE32(eocd + 12);
    uint16_t cdCount  = LE16(eocd + 10);

    std::vector<uint8_t> cd(cdSize);
    if (!ReadAt(m_hFile, cdOffset, cd.data(), cdSize)) return false;

    size_t pos = 0;
    for (uint16_t i = 0; i < cdCount; ++i) {
        if (pos + 46 > cdSize) break;
        if (LE32(cd.data() + pos) != 0x02014B50) break; // central dir sig

        uint16_t method    = LE16(cd.data() + pos + 10);
        uint32_t crc       = LE32(cd.data() + pos + 16);
        uint32_t compSz    = LE32(cd.data() + pos + 20);
        uint32_t uncompSz  = LE32(cd.data() + pos + 24);
        uint16_t nameLen   = LE16(cd.data() + pos + 28);
        uint16_t extraLen  = LE16(cd.data() + pos + 30);
        uint16_t commentLen= LE16(cd.data() + pos + 32);
        uint32_t localOff  = LE32(cd.data() + pos + 42);

        std::string name(reinterpret_cast<char*>(cd.data() + pos + 46), nameLen);

        ZipEntry entry;
        entry.name             = name;
        entry.method           = method;
        entry.crc32            = crc;
        entry.compressedSize   = compSz;
        entry.uncompressedSize = uncompSz;
        entry.offset           = localOff;
        m_entries[name]        = entry;

        pos += 46 + nameLen + extraLen + commentLen;
    }
    return !m_entries.empty();
}

bool ZipReader::HasEntry(const std::string& name) const {
    return m_entries.count(name) > 0;
}

std::vector<uint8_t> ZipReader::Extract(const std::string& name) {
    auto it = m_entries.find(name);
    if (it == m_entries.end()) return {};

    const ZipEntry& e = it->second;

    // Read local file header to get actual data offset
    uint8_t lfh[30];
    if (!ReadAt(m_hFile, e.offset, lfh, 30)) return {};
    if (LE32(lfh) != 0x04034B50) return {};
    uint16_t nameLen  = LE16(lfh + 26);
    uint16_t extraLen = LE16(lfh + 28);
    uint64_t dataOff  = e.offset + 30 + nameLen + extraLen;

    std::vector<uint8_t> compData(e.compressedSize);
    if (!ReadAt(m_hFile, dataOff, compData.data(), e.compressedSize)) return {};

    if (e.method == 0) {
        // Stored
        return compData;
    } else if (e.method == 8) {
        // Deflate
        std::vector<uint8_t> out(e.uncompressedSize);
        z_stream zs{};
        zs.next_in   = compData.data();
        zs.avail_in  = e.compressedSize;
        zs.next_out  = out.data();
        zs.avail_out = e.uncompressedSize;

        // -15: raw deflate (no zlib header)
        if (inflateInit2(&zs, -15) != Z_OK) return {};
        int ret = inflate(&zs, Z_FINISH);
        inflateEnd(&zs);
        if (ret != Z_STREAM_END) return {};
        out.resize(zs.total_out);
        return out;
    }
    return {}; // unsupported compression
}

std::string ZipReader::ExtractText(const std::string& name) {
    auto bytes = Extract(name);
    if (bytes.empty()) return {};
    return std::string(reinterpret_cast<char*>(bytes.data()), bytes.size());
}

std::wstring ZipReader::ExtractWide(const std::string& name) {
    std::string utf8 = ExtractText(name);
    if (utf8.empty()) return {};
    int needed = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(),
                                     static_cast<int>(utf8.size()), nullptr, 0);
    std::wstring ws(needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(),
                        static_cast<int>(utf8.size()), ws.data(), needed);
    return ws;
}

// --- ZipWriter ---

static void PutLE16(uint8_t* p, uint16_t v) {
    p[0] = v & 0xFF; p[1] = (v >> 8) & 0xFF;
}
static void PutLE32(uint8_t* p, uint32_t v) {
    p[0] = v & 0xFF; p[1] = (v >> 8) & 0xFF;
    p[2] = (v >> 16) & 0xFF; p[3] = (v >> 24) & 0xFF;
}

uint32_t ZipWriter::Crc32(const uint8_t* data, size_t size) {
    return static_cast<uint32_t>(::crc32(0, data, static_cast<uInt>(size)));
}

static bool WriteBytes(HANDLE h, const void* data, DWORD size) {
    DWORD written = 0;
    return WriteFile(h, data, size, &written, nullptr) && written == size;
}

static uint64_t CurrentOffset(HANDLE h) {
    LARGE_INTEGER li{}, result{};
    SetFilePointerEx(h, li, &result, FILE_CURRENT);
    return static_cast<uint64_t>(result.QuadPart);
}

bool ZipWriter::Open(const std::wstring& path) {
    m_hFile = CreateFileW(path.c_str(), GENERIC_WRITE, 0,
                          nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    return m_hFile != INVALID_HANDLE_VALUE;
}

bool ZipWriter::AddFile(const std::string& entryName,
                        const std::vector<uint8_t>& data, bool deflate) {
    LocalRecord rec;
    rec.name              = entryName;
    rec.uncompressedSize  = static_cast<uint32_t>(data.size());
    rec.crc32             = Crc32(data.data(), data.size());
    rec.localHeaderOffset = static_cast<uint32_t>(CurrentOffset(m_hFile));

    std::vector<uint8_t> compressed;
    if (deflate && !data.empty()) {
        // Compress using raw deflate
        uLongf bound = compressBound(static_cast<uLong>(data.size()));
        compressed.resize(bound);
        z_stream zs{};
        deflateInit2(&zs, Z_DEFAULT_COMPRESSION, Z_DEFLATED, -15, 8, Z_DEFAULT_STRATEGY);
        zs.next_in   = const_cast<Bytef*>(data.data());
        zs.avail_in  = static_cast<uInt>(data.size());
        zs.next_out  = compressed.data();
        zs.avail_out = static_cast<uInt>(bound);
        ::deflate(&zs, Z_FINISH);
        deflateEnd(&zs);
        compressed.resize(zs.total_out);
        rec.method         = 8;
        rec.compressedSize = static_cast<uint32_t>(compressed.size());
    } else {
        compressed         = data;
        rec.method         = 0;
        rec.compressedSize = rec.uncompressedSize;
    }

    return WriteLocalHeader(rec, compressed);
}

bool ZipWriter::AddText(const std::string& entryName, const std::string& utf8Text, bool def) {
    std::vector<uint8_t> data(utf8Text.begin(), utf8Text.end());
    return AddFile(entryName, data, def);
}

bool ZipWriter::WriteLocalHeader(const LocalRecord& rec, const std::vector<uint8_t>& compressed) {
    uint8_t lfh[30]{};
    PutLE32(lfh + 0, 0x04034B50); // signature
    PutLE16(lfh + 4, 20);         // version needed
    PutLE16(lfh + 6, 0);          // flags
    PutLE16(lfh + 8, static_cast<uint16_t>(rec.method));
    PutLE16(lfh + 10, 0);         // mod time
    PutLE16(lfh + 12, 0);         // mod date
    PutLE32(lfh + 14, rec.crc32);
    PutLE32(lfh + 18, rec.compressedSize);
    PutLE32(lfh + 22, rec.uncompressedSize);
    PutLE16(lfh + 26, static_cast<uint16_t>(rec.name.size()));
    PutLE16(lfh + 28, 0); // extra

    if (!WriteBytes(m_hFile, lfh, 30)) return false;
    if (!WriteBytes(m_hFile, rec.name.c_str(), static_cast<DWORD>(rec.name.size()))) return false;
    if (!compressed.empty())
        if (!WriteBytes(m_hFile, compressed.data(), static_cast<DWORD>(compressed.size()))) return false;

    m_records.push_back(rec);
    return true;
}

bool ZipWriter::Close() {
    if (m_hFile == INVALID_HANDLE_VALUE) return true;

    uint32_t cdOffset = static_cast<uint32_t>(CurrentOffset(m_hFile));

    for (const auto& rec : m_records) {
        uint8_t cdfh[46]{};
        PutLE32(cdfh + 0, 0x02014B50); // central dir sig
        PutLE16(cdfh + 4, 20);         // version made by
        PutLE16(cdfh + 6, 20);         // version needed
        PutLE16(cdfh + 8, 0);          // flags
        PutLE16(cdfh + 10, static_cast<uint16_t>(rec.method));
        PutLE16(cdfh + 12, 0); // time
        PutLE16(cdfh + 14, 0); // date
        PutLE32(cdfh + 16, rec.crc32);
        PutLE32(cdfh + 20, rec.compressedSize);
        PutLE32(cdfh + 24, rec.uncompressedSize);
        PutLE16(cdfh + 28, static_cast<uint16_t>(rec.name.size()));
        PutLE16(cdfh + 30, 0); // extra
        PutLE16(cdfh + 32, 0); // comment
        PutLE16(cdfh + 34, 0); // disk start
        PutLE16(cdfh + 36, 0); // int attrs
        PutLE32(cdfh + 38, 0); // ext attrs
        PutLE32(cdfh + 42, rec.localHeaderOffset);

        WriteBytes(m_hFile, cdfh, 46);
        WriteBytes(m_hFile, rec.name.c_str(), static_cast<DWORD>(rec.name.size()));
    }

    uint32_t cdEnd = static_cast<uint32_t>(CurrentOffset(m_hFile));
    uint32_t cdSize = cdEnd - cdOffset;

    uint8_t eocd[22]{};
    PutLE32(eocd + 0, 0x06054B50);  // EOCD sig
    PutLE16(eocd + 4, 0);           // disk number
    PutLE16(eocd + 6, 0);           // disk with CD start
    PutLE16(eocd + 8, static_cast<uint16_t>(m_records.size()));
    PutLE16(eocd + 10, static_cast<uint16_t>(m_records.size()));
    PutLE32(eocd + 12, cdSize);
    PutLE32(eocd + 16, cdOffset);
    PutLE16(eocd + 20, 0); // comment length
    WriteBytes(m_hFile, eocd, 22);

    CloseHandle(m_hFile);
    m_hFile = INVALID_HANDLE_VALUE;
    m_records.clear();
    return true;
}
