#pragma once
#include "IDocumentFormat.h"

class MdFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    // Convert Markdown source to RTF (headings, bold, italic, code)
    static std::string MdToRtf(const std::wstring& md);
    // Convert inline Markdown spans to RTF
    static std::string InlineToRtf(const std::wstring& line);
};
