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
    // Convert Markdown to plain readable text (strip formatting markers)
    static std::wstring MdToPlain(const std::wstring& md);
};
