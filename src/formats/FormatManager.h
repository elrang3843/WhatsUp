#pragma once
#include "IDocumentFormat.h"
#include "../Document.h"
#include <memory>

class FormatManager {
public:
    static FormatManager& Instance();

    FormatResult Load(const std::wstring& path, Document& doc);
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document&           doc);

private:
    FormatManager() = default;
    std::unique_ptr<IDocumentFormat> HandlerFor(DocFormat fmt);
};
