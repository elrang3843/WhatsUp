#pragma once
#include "IDocumentFormat.h"
#include "../cdm/document_model.hpp"
#include <map>

class DocxFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc) override;
    bool CanWrite() const override { return true; }

private:
    // Convert word/document.xml to CDM document. `imgResources` maps a
    // relationship id (rIdN referenced by <a:blip r:embed>) to a fully
    // populated cdm::Resource (bytes + mediaType); the parser uses this to
    // emit real ResourceIds on Image nodes and to register resources on the
    // Document it builds.
    static cdm::Document ParseDocumentXmlToCdm(
        const std::wstring& xml,
        const std::map<std::wstring, cdm::ListType>& numTypes,
        const std::map<std::wstring, std::wstring>& rels,
        const std::map<std::wstring, cdm::Resource>& imgResources);
    static std::map<std::wstring, cdm::ListType> ParseNumbering(const std::wstring& xml);
    static std::map<std::wstring, std::wstring>  ParseDocRels(const std::wstring& xml);
    // Build word/document.xml from plain text
    static std::string  BuildDocumentXml(const std::wstring& text, const DocProperties& props);
    static std::string  BuildCoreXml(const DocProperties& props);
    static std::string  BuildContentTypes();
    static std::string  BuildRels();
    static std::string  BuildWordRels();
    static std::string  BuildSettings();
    static std::string  BuildStyles();
};
