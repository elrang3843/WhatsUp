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
    static std::wstring WrapHtml(const std::wstring& text, const DocProperties& props);
};
