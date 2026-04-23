#pragma once
#include "IDocumentFormat.h"
#include "../cdm/document_model.hpp"

class DocxFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    // Convert word/document.xml to CDM document
    static cdm::Document ParseDocumentXmlToCdm(const std::wstring& xml);
    // Build word/document.xml from plain text
    static std::string  BuildDocumentXml(const std::wstring& text, const DocProperties& props);
    static std::string  BuildCoreXml(const DocProperties& props);
    static std::string  BuildContentTypes();
    static std::string  BuildRels();
    static std::string  BuildWordRels();
    static std::string  BuildSettings();
    static std::string  BuildStyles();
};
