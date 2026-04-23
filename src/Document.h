#pragma once
#include <windows.h>
#include <string>

enum class DocFormat {
    Unknown,
    Txt,
    Html,
    Md,
    Docx,
    Hwpx,
    Hwp,
    Doc,
};

enum class WatermarkType { None, Text, Image };

struct WatermarkInfo {
    WatermarkType type    = WatermarkType::None;
    std::wstring  text;
    std::wstring  imagePath;
    int           opacity = 30;   // 0–100
    int           angle   = 45;
    COLORREF      color   = RGB(192, 192, 192);
};

struct DocProperties {
    std::wstring title;
    std::wstring author;
    std::wstring subject;
    std::wstring keywords;
    std::wstring comment;
    WatermarkInfo watermark;
};

struct DocStats {
    int words = 0;
    int chars = 0;
    int pages = 1;
};

class Document {
public:
    Document();

    // File path management
    const std::wstring& Path() const    { return m_path; }
    void  SetPath(const std::wstring& p){ m_path = p; }
    bool  HasPath() const               { return !m_path.empty(); }

    // Format detection
    DocFormat  Format() const            { return m_format; }
    void       SetFormat(DocFormat f)    { m_format = f; }
    bool       IsReadOnly() const        { return m_readOnly; }
    void       SetReadOnly(bool v)       { m_readOnly = v; }

    // Modified flag
    bool  IsModified() const             { return m_modified; }
    void  SetModified(bool v)            { m_modified = v; }

    // Properties
    DocProperties&       Properties()       { return m_props; }
    const DocProperties& Properties() const { return m_props; }

    // Statistics (updated by MainWindow after each text change)
    DocStats&       Stats()       { return m_stats; }
    const DocStats& Stats() const { return m_stats; }

    // Title for titlebar ("제목 없음" or filename, with * if modified)
    std::wstring DisplayTitle() const;

    // Detect format from file extension
    static DocFormat DetectFormat(const std::wstring& path);

    // Human-readable format name
    static const wchar_t* FormatName(DocFormat f);

    void Reset();

private:
    std::wstring  m_path;
    DocFormat     m_format   = DocFormat::Txt;
    bool          m_modified = false;
    bool          m_readOnly = false;
    DocProperties m_props;
    DocStats      m_stats;
};
