#pragma once
#include "IDocumentFormat.h"

// HWP 5.x binary format — read-only text extraction
// Full HWP support requires the HWP5 SDK or reverse-engineered spec
class HwpFormat : public IDocumentFormat {
public:
    FormatResult Load(const std::wstring& path, Document& doc) override;
    FormatResult Save(const std::wstring& path,
                      const std::wstring& content,
                      const std::string&  rtfContent,
                      Document& doc) override;
    bool CanWrite() const override { return false; }

private:
    // Attempt to extract readable Unicode text from HWP5 binary streams
    static std::wstring ExtractText(const std::vector<uint8_t>& data);
    // Read an OLE2 Compound Document stream by name
    static std::vector<uint8_t> ReadOleStream(const std::wstring& path,
                                              const std::wstring& streamName);
};
