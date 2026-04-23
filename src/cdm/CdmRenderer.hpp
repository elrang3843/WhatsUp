#pragma once
#include "document_model.hpp"
#include <string>

namespace cdm {

// Converts a cdm::Document to an RTF string suitable for EM_STREAMIN into a RichEdit control.
class CdmRenderer {
public:
    std::string Render(const Document& doc);

private:
    std::string rtf_;

    // two-pass color/font collection
    std::vector<uint32_t>    colorTable_;  // 0xRRGGBB
    std::vector<std::string> fontTable_;

    int indentLeft_ = 0; // accumulated left indent in twips (for blockquotes)

    int  ColorIndex(const Color& c);
    int  FontIndex(const std::string& family);

    void BuildTables(const Document& doc);
    void CollectBlock(const Block& blk);
    void CollectInline(const Inline& inl);

    void WriteHeader();
    void WriteDocument(const Document& doc);
    void WriteSection(const Section& sec);
    void WriteBlock(const Block& blk);
    void WriteParagraph(const Paragraph& p, bool isList = false);
    void WriteHeading(const Heading& h);
    void WriteCodeBlock(const CodeBlock& cb);
    void WriteHRule();
    void WriteTable(const Table& t);
    void WriteBlockQuote(const BlockQuote& bq);
    void WriteList(const ListBlock& lb);
    void WriteInlines(const std::vector<InlinePtr>& inlines);
    void WriteInline(const Inline& inl);
    void WriteText(const Text& t);
    void WriteInlineContainer(const InlineContainer& ic, const char* onCmd, const char* offCmd);

    void ApplyTextStyle(const TextStyle& s, std::string& out);
    void ResetStyle(std::string& out);

    static std::string EscapeRtf(const std::u32string& text);
    static std::string EscapeRtfA(const std::string& text);
};

} // namespace cdm
