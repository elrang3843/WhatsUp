#pragma once
#include "IDocumentFormat.h"

class DocxFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    // Extract all text runs from word/document.xml
    static std::wstring ParseDocumentXml(const std::wstring& xml);
    // Build word/document.xml from plain text
    static std::string  BuildDocumentXml(const std::wstring& text, const DocProperties& props);
    // Build minimal docProps/core.xml
    static std::string  BuildCoreXml(const DocProperties& props);
    // Build [Content_Types].xml
    static std::string  BuildContentTypes();
    // Build _rels/.rels
    static std::string  BuildRels();
    // Build word/_rels/document.xml.rels
    static std::string  BuildWordRels();
    // Build word/settings.xml
    static std::string  BuildSettings();
    // Build word/styles.xml
    static std::string  BuildStyles();
};
