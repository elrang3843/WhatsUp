#pragma once
#include "IDocumentFormat.h"

// HWPX is an XML-based format (ZIP container), standardized as ISO/IEC 26300-3
class HwpxFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    static std::wstring ParseBodyText(const std::wstring& xml);
    static std::string  BuildBodyText(const std::wstring& text);
    static std::string  BuildContentType();
    static std::string  BuildRels();
    static std::string  BuildManifest();
    static std::string  BuildSummary(const DocProperties& props);
};
