#pragma once
#include <windows.h>
#include <string>
#include "../Document.h"
#include "../cdm/document_model.hpp"

// Result returned by format handlers
struct FormatResult {
    bool            ok      = false;
    std::wstring    error;
    std::wstring    content; // legacy: plain/RTF text (UTF-16)
    bool            rtf     = false; // true if content is RTF, false if plain Unicode
    cdm::Document   cdmDoc;          // CDM representation (non-empty sections = CDM path)
};

// Abstract interface for all document format handlers
class IDocumentFormat {
public:
    virtual ~IDocumentFormat() = default;

    virtual FormatResult Load(const std::wstring& path, Document& doc) = 0;

    virtual FormatResult Save(const std::wstring& path,
                              const std::wstring& content,
                              const std::string&  rtfContent,
                              Document&           doc) = 0;

    virtual bool CanWrite() const = 0;
};
