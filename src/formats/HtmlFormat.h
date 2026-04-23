#pragma once
#include "IDocumentFormat.h"

class HtmlFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    // Strip HTML tags to extract plain text
    static std::wstring StripTags(const std::wstring& html);
    // Convert &amp; &lt; &gt; &nbsp; etc.
    static std::wstring DecodeEntities(const std::wstring& text);
    // Build a minimal HTML wrapper around plain text
    static std::wstring WrapHtml(const std::wstring& text, const DocProperties& props);
};
