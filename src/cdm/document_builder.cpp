#include "document_builder.hpp"
#include <stdexcept>

namespace cdm {

// ---- ctor ----

DocumentBuilder::DocumentBuilder() {
    doc_.id = idGen_.Next();
}

// ---- build ----

Document DocumentBuilder::Build() const { return doc_; }
Document&& DocumentBuilder::MoveBuild()  { return std::move(doc_); }

// ---- metadata ----

DocumentBuilder& DocumentBuilder::SetOriginalFormat(FileFormat fmt) {
    doc_.originalFormat = fmt; return *this;
}
DocumentBuilder& DocumentBuilder::SetTitle(const std::string& v)       { doc_.properties.title       = v; return *this; }
DocumentBuilder& DocumentBuilder::SetSubject(const std::string& v)     { doc_.properties.subject     = v; return *this; }
DocumentBuilder& DocumentBuilder::SetAuthor(const std::string& v)      { doc_.properties.author      = v; return *this; }
DocumentBuilder& DocumentBuilder::SetCreator(const std::string& v)     { doc_.properties.creator     = v; return *this; }
DocumentBuilder& DocumentBuilder::SetKeywords(const std::string& v)    { doc_.properties.keywords    = v; return *this; }
DocumentBuilder& DocumentBuilder::SetDescription(const std::string& v) { doc_.properties.description = v; return *this; }
DocumentBuilder& DocumentBuilder::SetCreatedAt(const std::string& v)   { doc_.properties.createdAt   = v; return *this; }
DocumentBuilder& DocumentBuilder::SetModifiedAt(const std::string& v)  { doc_.properties.modifiedAt  = v; return *this; }

DocumentBuilder& DocumentBuilder::SetDefaultLanguage(const std::string& bcp47) {
    doc_.textEnvironment.defaultLanguage = LanguageTag{bcp47}; return *this;
}
DocumentBuilder& DocumentBuilder::SetDefaultScript(Script s) {
    doc_.textEnvironment.defaultScript = s; return *this;
}
DocumentBuilder& DocumentBuilder::SetDefaultCharset(FontCharset c) {
    doc_.textEnvironment.defaultCharset = c; return *this;
}

EncodingInfo DocumentBuilder::MakeEncodingInfo(TextEncoding enc, const std::string& name,
    std::optional<uint32_t> cp, bool bom, bool declared, bool heuristic) const
{
    EncodingInfo ei;
    ei.encoding              = enc;
    ei.originalName          = name;
    ei.codePage              = cp;
    ei.hadBOM                = bom;
    ei.declaredInDocument    = declared;
    ei.detectedHeuristically = heuristic;
    return ei;
}

DocumentBuilder& DocumentBuilder::SetSourceEncoding(TextEncoding enc,
    const std::string& name, std::optional<uint32_t> cp,
    bool bom, bool declared, bool heuristic)
{
    doc_.textEnvironment.sourceEncoding =
        MakeEncodingInfo(enc, name, cp, bom, declared, heuristic);
    return *this;
}

DocumentBuilder& DocumentBuilder::SetPreferredSaveEncoding(TextEncoding enc,
    const std::string& name, std::optional<uint32_t> cp, bool bom)
{
    doc_.textEnvironment.preferredSaveEncoding =
        MakeEncodingInfo(enc, name, cp, bom, false, false);
    return *this;
}

DocumentBuilder& DocumentBuilder::SetMetadata(const std::string& key, const std::string& value) {
    doc_.metadata[key] = value; return *this;
}

// ---- collections ----

DocumentBuilder& DocumentBuilder::AddStyle(const StyleDefinition& style) {
    doc_.styles.push_back(style); return *this;
}
DocumentBuilder& DocumentBuilder::AddResource(const Resource& resource) {
    doc_.resources.push_back(resource); return *this;
}
DocumentBuilder& DocumentBuilder::AddComment(const Comment& comment) {
    doc_.comments.push_back(comment); return *this;
}
DocumentBuilder& DocumentBuilder::AddNote(const Note& note) {
    doc_.notes.push_back(note); return *this;
}
DocumentBuilder& DocumentBuilder::AddBookmark(const Bookmark& bm) {
    doc_.bookmarks.push_back(bm); return *this;
}

// ---- section ----

SectionPtr DocumentBuilder::EnsureSection() {
    if (!currentSection_) {
        currentSection_ = MakeNode<Section>(idGen_);
        doc_.sections.push_back(currentSection_);
    }
    return currentSection_;
}

DocumentBuilder& DocumentBuilder::BeginSection() {
    currentSection_ = MakeNode<Section>(idGen_);
    doc_.sections.push_back(currentSection_);
    return *this;
}

DocumentBuilder& DocumentBuilder::EndSection() {
    currentSection_ = nullptr;
    return *this;
}

// ---- helpers ----

void DocumentBuilder::EnsureNoOpenParagraph() const {
    if (currentParagraph_)
        throw std::logic_error("DocumentBuilder: paragraph already open");
}

void DocumentBuilder::RequireOpenParagraph() const {
    if (!currentParagraph_)
        throw std::logic_error("DocumentBuilder: no open paragraph");
}

std::vector<BlockPtr>& DocumentBuilder::CurrentBlockContainer() {
    if (currentTableRow_) {
        // inside a table row — should not add blocks directly
        throw std::logic_error("DocumentBuilder: cannot add block inside table row directly");
    }
    if (currentList_)
        return currentList_->items.empty()
            ? EnsureSection()->blocks
            : currentList_->items.back()->blocks;
    if (currentBlockQuote_)
        return currentBlockQuote_->blocks;
    return EnsureSection()->blocks;
}

void DocumentBuilder::PushBlockToCurrentContainer(BlockPtr block) {
    CurrentBlockContainer().push_back(std::move(block));
}

std::vector<InlinePtr>& DocumentBuilder::CurrentInlineContainer() {
    RequireOpenParagraph();
    if (!inlineCtxStack_.empty())
        return *inlineCtxStack_.back().inlines;
    return currentParagraph_->inlines;
}

// ---- inline factory helpers ----

InlinePtr DocumentBuilder::MakeTextInline(const std::u32string& text) {
    auto inl = MakeInline<Text>(idGen_);
    std::get<Text>(inl->value).text = text;
    return inl;
}

BlockPtr DocumentBuilder::MakeParagraphBlock(const std::u32string& text) {
    auto blk = MakeBlock<Paragraph>(idGen_);
    if (!text.empty()) {
        auto& para = std::get<Paragraph>(blk->value);
        para.inlines.push_back(MakeTextInline(text));
    }
    return blk;
}

BlockPtr DocumentBuilder::MakeHeadingBlock(int level, const std::u32string& text) {
    auto blk = MakeBlock<Heading>(idGen_);
    auto& h   = std::get<Heading>(blk->value);
    h.level   = level;
    if (!text.empty())
        h.inlines.push_back(MakeTextInline(text));
    return blk;
}

// ---- blocks ----

DocumentBuilder& DocumentBuilder::AddHeading(int level, const std::u32string& text) {
    EnsureNoOpenParagraph();
    PushBlockToCurrentContainer(MakeHeadingBlock(level, text));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddParagraph(const std::u32string& text) {
    EnsureNoOpenParagraph();
    PushBlockToCurrentContainer(MakeParagraphBlock(text));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddCodeBlock(const std::string& code, const std::string& lang) {
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<CodeBlock>(idGen_);
    auto& cb = std::get<CodeBlock>(blk->value);
    cb.code     = code;
    cb.language = lang;
    PushBlockToCurrentContainer(std::move(blk));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddHorizontalRule() {
    EnsureNoOpenParagraph();
    PushBlockToCurrentContainer(MakeBlock<HorizontalRule>(idGen_));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddRawBlock(RawFragmentKind kind,
    const std::string& data, std::optional<EncodingInfo> enc)
{
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<RawBlock>(idGen_);
    auto& rb = std::get<RawBlock>(blk->value);
    rb.kind     = kind;
    rb.data     = data;
    rb.encoding = enc;
    PushBlockToCurrentContainer(std::move(blk));
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginParagraph() {
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<Paragraph>(idGen_);
    // Keep pointer into the block before ownership is transferred
    currentParagraph_ = &std::get<Paragraph>(blk->value);
    PushBlockToCurrentContainer(std::move(blk));
    return *this;
}

DocumentBuilder& DocumentBuilder::EndParagraph() {
    RequireOpenParagraph();
    inlineCtxStack_.clear();
    currentParagraph_ = nullptr;
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginBlockQuote() {
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<BlockQuote>(idGen_);
    currentBlockQuote_ = &std::get<BlockQuote>(blk->value);
    PushBlockToCurrentContainer(std::move(blk));
    contextStack_.push_back(OpenBlockContext::BlockQuote);
    return *this;
}

DocumentBuilder& DocumentBuilder::EndBlockQuote() {
    currentBlockQuote_ = nullptr;
    if (!contextStack_.empty()) contextStack_.pop_back();
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginList(ListType type, int32_t start, int32_t level) {
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<ListBlock>(idGen_);
    auto& lb = std::get<ListBlock>(blk->value);
    lb.type  = type;
    lb.start = start;
    lb.level = level;
    currentList_ = &lb;
    PushBlockToCurrentContainer(std::move(blk));
    contextStack_.push_back(OpenBlockContext::List);
    return *this;
}

DocumentBuilder& DocumentBuilder::AddListItem(const std::u32string& text) {
    if (!currentList_)
        throw std::logic_error("DocumentBuilder: no open list");
    auto item = MakeNode<ListItem>(idGen_);
    item->blocks.push_back(MakeParagraphBlock(text));
    currentList_->items.push_back(std::move(item));
    return *this;
}

DocumentBuilder& DocumentBuilder::EndList() {
    currentList_ = nullptr;
    if (!contextStack_.empty()) contextStack_.pop_back();
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginTable() {
    EnsureNoOpenParagraph();
    auto blk = MakeBlock<Table>(idGen_);
    currentTable_ = &std::get<Table>(blk->value);
    PushBlockToCurrentContainer(std::move(blk));
    contextStack_.push_back(OpenBlockContext::Table);
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginTableRow() {
    if (!currentTable_)
        throw std::logic_error("DocumentBuilder: no open table");
    auto row = MakeNode<TableRow>(idGen_);
    currentTableRow_ = row.get();
    currentTable_->rows.push_back(std::move(row));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddTableCell(const std::u32string& text) {
    if (!currentTableRow_)
        throw std::logic_error("DocumentBuilder: no open table row");
    auto cell = MakeNode<TableCell>(idGen_);
    cell->blocks.push_back(MakeParagraphBlock(text));
    currentTableRow_->cells.push_back(std::move(cell));
    return *this;
}

DocumentBuilder& DocumentBuilder::EndTableRow() {
    currentTableRow_ = nullptr;
    return *this;
}

DocumentBuilder& DocumentBuilder::EndTable() {
    currentTable_    = nullptr;
    currentTableRow_ = nullptr;
    if (!contextStack_.empty()) contextStack_.pop_back();
    return *this;
}

// ---- inlines ----

DocumentBuilder& DocumentBuilder::AddText(const std::u32string& text) {
    CurrentInlineContainer().push_back(MakeTextInline(text));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddStyledText(const std::u32string& text, const TextStyle& style) {
    auto inl = MakeInline<Text>(idGen_);
    auto& t  = std::get<Text>(inl->value);
    t.text        = text;
    t.directStyle = style;
    CurrentInlineContainer().push_back(std::move(inl));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddTab() {
    CurrentInlineContainer().push_back(MakeInline<Tab>(idGen_));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddLineBreak() {
    auto inl = MakeInline<Break>(idGen_);
    std::get<Break>(inl->value).type = BreakType::LineBreak;
    CurrentInlineContainer().push_back(std::move(inl));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddSoftBreak() {
    auto inl = MakeInline<Break>(idGen_);
    std::get<Break>(inl->value).type = BreakType::SoftBreak;
    CurrentInlineContainer().push_back(std::move(inl));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddLink(const std::string& href,
    const std::string& displayText, const std::string& title)
{
    // Wrap link text in a Hyperlink span with children
    auto linkInl = MakeInline<Hyperlink>(idGen_);
    auto& lnk    = std::get<Hyperlink>(linkInl->value);
    lnk.href  = href;
    lnk.title = title;
    (void)displayText; // Hyperlink is inline-container-less in the model; add text as sibling
    // Per the model, Hyperlink has no children vector — emit display text as text node instead
    if (!displayText.empty()) {
        std::u32string u32;
        for (unsigned char c : displayText) u32 += static_cast<char32_t>(c);
        CurrentInlineContainer().push_back(MakeTextInline(u32));
    }
    CurrentInlineContainer().push_back(std::move(linkInl));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddImage(ResourceId rid,
    const std::string& alt, const std::string& title)
{
    auto inl = MakeInline<Image>(idGen_);
    auto& img = std::get<Image>(inl->value);
    img.resourceId = rid;
    img.altText    = alt;
    img.title      = title;
    CurrentInlineContainer().push_back(std::move(inl));
    return *this;
}

DocumentBuilder& DocumentBuilder::AddRawInline(RawFragmentKind kind,
    const std::string& data, std::optional<EncodingInfo> enc)
{
    auto inl = MakeInline<RawInline>(idGen_);
    auto& ri = std::get<RawInline>(inl->value);
    ri.kind     = kind;
    ri.data     = data;
    ri.encoding = enc;
    CurrentInlineContainer().push_back(std::move(inl));
    return *this;
}

// ---- inline containers ----

template<typename ContainerT>
DocumentBuilder& DocumentBuilder::BeginInlineContainer(InlineCtxType ctxType) {
    RequireOpenParagraph();
    auto inl = MakeInline<ContainerT>(idGen_);
    auto& container = std::get<ContainerT>(inl->value).children;
    inlineCtxStack_.push_back({ ctxType, inl, &container });
    // The node itself will be appended to the parent when End* is called
    return *this;
}

template<typename ContainerT>
DocumentBuilder& DocumentBuilder::EndInlineContainer(InlineCtxType ctxType, const char* name) {
    if (inlineCtxStack_.empty() || inlineCtxStack_.back().type != ctxType)
        throw std::logic_error(std::string("DocumentBuilder: mismatched End") + name);
    InlineCtx ctx = std::move(inlineCtxStack_.back());
    inlineCtxStack_.pop_back();
    // Append to parent inline container (or paragraph)
    std::vector<InlinePtr>& parent = inlineCtxStack_.empty()
        ? currentParagraph_->inlines
        : *inlineCtxStack_.back().inlines;
    parent.push_back(std::move(ctx.node));
    return *this;
}

DocumentBuilder& DocumentBuilder::BeginStrong()   { return BeginInlineContainer<Strong>(InlineCtxType::Strong); }
DocumentBuilder& DocumentBuilder::EndStrong()     { return EndInlineContainer<Strong>(InlineCtxType::Strong,   "Strong"); }
DocumentBuilder& DocumentBuilder::BeginEmphasis() { return BeginInlineContainer<Emphasis>(InlineCtxType::Emphasis); }
DocumentBuilder& DocumentBuilder::EndEmphasis()   { return EndInlineContainer<Emphasis>(InlineCtxType::Emphasis, "Emphasis"); }
DocumentBuilder& DocumentBuilder::BeginUnderline(){ return BeginInlineContainer<Underline>(InlineCtxType::Underline); }
DocumentBuilder& DocumentBuilder::EndUnderline()  { return EndInlineContainer<Underline>(InlineCtxType::Underline,"Underline"); }

DocumentBuilder& DocumentBuilder::BeginSpan(const TextStyle& style) {
    RequireOpenParagraph();
    auto inl = MakeInline<Span>(idGen_);
    std::get<Span>(inl->value).directStyle = style;
    auto& children = std::get<Span>(inl->value).children;
    inlineCtxStack_.push_back({ InlineCtxType::Span, inl, &children });
    return *this;
}
DocumentBuilder& DocumentBuilder::EndSpan() { return EndInlineContainer<Span>(InlineCtxType::Span, "Span"); }

// ---- paragraph styling ----

DocumentBuilder& DocumentBuilder::SetCurrentParagraphStyleRef(const StyleId& styleId) {
    RequireOpenParagraph();
    currentParagraph_->styleRef = styleId;
    return *this;
}

DocumentBuilder& DocumentBuilder::SetCurrentParagraphLanguage(const std::string& bcp47) {
    RequireOpenParagraph();
    currentParagraph_->directStyle.language = LanguageTag{bcp47};
    return *this;
}

DocumentBuilder& DocumentBuilder::SetCurrentParagraphAlignment(Alignment align) {
    RequireOpenParagraph();
    currentParagraph_->directStyle.alignment = align;
    return *this;
}

} // namespace cdm
