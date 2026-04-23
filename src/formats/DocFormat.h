#pragma once
#include "IDocumentFormat.h"

// DOC (Word 97-2003 binary) — read-only plain text extraction
// Full DOC support requires implementing the MS-DOC specification (OLE2 + Word Binary)
class DocFormat_ : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document& doc) override;
    bool CanWrite() const override { return false; }

private:
    // Extract text from the WordDocument stream's text array (FIB + PlcfBte approach)
    static std::wstring ExtractText(const std::wstring& path);
};
