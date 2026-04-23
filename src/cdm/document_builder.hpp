#pragma once
#include "document_model.hpp"
#include <stdexcept>
#include <string>
#include <utility>

namespace cdm {

class DocumentBuilder {
public:
    DocumentBuilder();

    Document  Build()     const;
    Document&& MoveBuild();

    // ---- Metadata ----
    DocumentBuilder& SetOriginalFormat(FileFormat fmt);
    DocumentBuilder& SetTitle(const std::string& v);
    DocumentBuilder& SetSubject(const std::string& v);
    DocumentBuilder& SetAuthor(const std::string& v);
    DocumentBuilder& SetCreator(const std::string& v);
    DocumentBuilder& SetKeywords(const std::string& v);
    DocumentBuilder& SetDescription(const std::string& v);
    DocumentBuilder& SetCreatedAt(const std::string& v);
    DocumentBuilder& SetModifiedAt(const std::string& v);
    DocumentBuilder& SetDefaultLanguage(const std::string& bcp47);
    DocumentBuilder& SetDefaultScript(Script script);
    DocumentBuilder& SetDefaultCharset(FontCharset charset);
    DocumentBuilder& SetSourceEncoding(TextEncoding encoding,
        const std::string& originalName = {}, std::optional<uint32_t> codePage = {},
        bool hadBOM = false, bool declaredInDocument = false, bool detectedHeuristically = false);
    DocumentBuilder& SetPreferredSaveEncoding(TextEncoding encoding,
        const std::string& originalName = {}, std::optional<uint32_t> codePage = {}, bool hadBOM = false);
    DocumentBuilder& SetMetadata(const std::string& key, const std::string& value);

    // ---- Collections ----
    DocumentBuilder& AddStyle(const StyleDefinition& style);
    DocumentBuilder& AddResource(const Resource& resource);
    DocumentBuilder& AddComment(const Comment& comment);
    DocumentBuilder& AddNote(const Note& note);
    DocumentBuilder& AddBookmark(const Bookmark& bookmark);

    // ---- Section ----
    DocumentBuilder& BeginSection();
    DocumentBuilder& EndSection();

    // ---- Blocks ----
    DocumentBuilder& AddHeading(int level, const std::u32string& text);
    DocumentBuilder& AddParagraph(const std::u32string& text);
    DocumentBuilder& AddCodeBlock(const std::string& code, const std::string& language = {});
    DocumentBuilder& AddHorizontalRule();
    DocumentBuilder& AddRawBlock(RawFragmentKind kind, const std::string& data,
                                  std::optional<EncodingInfo> encoding = {});

    DocumentBuilder& BeginParagraph();
    DocumentBuilder& EndParagraph();
    DocumentBuilder& BeginBlockQuote();
    DocumentBuilder& EndBlockQuote();
    DocumentBuilder& BeginList(ListType type = ListType::Bullet, int32_t start = 1, int32_t level = 0);
    DocumentBuilder& AddListItem(const std::u32string& text);
    DocumentBuilder& EndList();
    DocumentBuilder& BeginTable();
    DocumentBuilder& BeginTableRow();
    DocumentBuilder& AddTableCell(const std::u32string& text);
    DocumentBuilder& EndTableRow();
    DocumentBuilder& EndTable();

    // ---- Inlines (valid inside BeginParagraph / EndParagraph) ----
    DocumentBuilder& AddText(const std::u32string& text);
    DocumentBuilder& AddStyledText(const std::u32string& text, const TextStyle& style);
    DocumentBuilder& AddTab();
    DocumentBuilder& AddLineBreak();
    DocumentBuilder& AddSoftBreak();
    DocumentBuilder& AddLink(const std::string& href, const std::string& displayText,
                              const std::string& title = {});
    DocumentBuilder& AddImage(ResourceId resourceId, const std::string& altText = {},
                               const std::string& title = {});
    DocumentBuilder& AddRawInline(RawFragmentKind kind, const std::string& data,
                                   std::optional<EncodingInfo> encoding = {});

    // ---- Inline containers (nesting supported) ----
    DocumentBuilder& BeginStrong();
    DocumentBuilder& EndStrong();
    DocumentBuilder& BeginEmphasis();
    DocumentBuilder& EndEmphasis();
    DocumentBuilder& BeginUnderline();
    DocumentBuilder& EndUnderline();
    DocumentBuilder& BeginSpan(const TextStyle& style = {});
    DocumentBuilder& EndSpan();

    // ---- Current paragraph styling ----
    DocumentBuilder& SetCurrentParagraphStyleRef(const StyleId& styleId);
    DocumentBuilder& SetCurrentParagraphLanguage(const std::string& bcp47);
    DocumentBuilder& SetCurrentParagraphAlignment(Alignment align);

private:
    enum class OpenBlockContext { None, Paragraph, BlockQuote, List, Table, TableRow };
    enum class InlineCtxType    { Strong, Emphasis, Underline, Span };

    struct InlineCtx {
        InlineCtxType           type;
        InlinePtr               node;        // the container Inline node
        std::vector<InlinePtr>* inlines;     // points into node's children
    };

    Document         doc_;
    NodeIdGenerator  idGen_;

    SectionPtr    currentSection_;
    Paragraph*    currentParagraph_   = nullptr;
    BlockQuote*   currentBlockQuote_  = nullptr;
    ListBlock*    currentList_        = nullptr;
    Table*        currentTable_       = nullptr;
    TableRow*     currentTableRow_    = nullptr;

    std::vector<OpenBlockContext> contextStack_;
    std::vector<InlineCtx>       inlineCtxStack_;

    SectionPtr           EnsureSection();
    void                 EnsureNoOpenParagraph() const;
    void                 RequireOpenParagraph()  const;
    void                 PushBlockToCurrentContainer(BlockPtr block);
    std::vector<BlockPtr>&  CurrentBlockContainer();
    std::vector<InlinePtr>& CurrentInlineContainer();

    EncodingInfo MakeEncodingInfo(TextEncoding enc, const std::string& name,
        std::optional<uint32_t> cp, bool bom, bool declared, bool heuristic) const;

    InlinePtr MakeTextInline(const std::u32string& text);
    BlockPtr  MakeParagraphBlock(const std::u32string& text);
    BlockPtr  MakeHeadingBlock(int level, const std::u32string& text);

    template<typename ContainerT>
    DocumentBuilder& BeginInlineContainer(InlineCtxType ctxType);
    template<typename ContainerT>
    DocumentBuilder& EndInlineContainer(InlineCtxType ctxType, const char* name);
};

} // namespace cdm
