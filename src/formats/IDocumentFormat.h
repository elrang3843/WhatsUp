#pragma once
#include <windows.h>
#include <string>
#include "../Document.h"

// Result returned by format handlers
struct FormatResult {
    bool         ok      = false;
    std::wstring error;
    std::wstring content; // extracted plain/rich text (UTF-16)
    bool         rtf     = false; // true if content is RTF, false if plain Unicode
};

// Abstract interface for all document format handlers
class IDocumentFormat {
public:
    virtual ~IDocumentFormat() = default;

    // Load the file at `path` into `content` (UTF-16 or RTF)
    virtual FormatResult Load(const std::wstring& path, Document& doc) = 0;

    // Save `content` (UTF-16 plain text) to `path`
    // For formats that support rich text, richContent (RTF bytes) is also provided
    virtual FormatResult Save(const std::wstring& path,
                              const std::wstring& content,
                              const std::string&  rtfContent,
                              Document&           doc) = 0;

    // Returns true if this handler can write (not just read)
    virtual bool CanWrite() const = 0;
};
