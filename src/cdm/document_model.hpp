#pragma once

#include <cstdint>
#include <map>
#include <memory>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <variant>
#include <vector>

namespace cdm {

using NodeId      = std::uint64_t;
using ResourceId  = std::uint64_t;
using StyleId     = std::string;
using MetadataMap = std::unordered_map<std::string, std::string>;
using BinaryBlob  = std::vector<std::uint8_t>;

enum class FileFormat { Unknown, HWP, HWPX, DOC, DOCX, HTM, HTML, MD, TXT };

enum class TextEncoding {
    Unknown, UTF8, UTF8_BOM, UTF16LE, UTF16BE, UTF32LE, UTF32BE,
    ASCII, CP949, EUC_KR, ShiftJIS, GB18030, Big5, Windows1252, ISO8859_1, Custom
};

struct EncodingInfo {
    TextEncoding             encoding              = TextEncoding::Unknown;
    std::string              originalName;
    std::optional<uint32_t>  codePage;
    bool                     hadBOM                = false;
    bool                     declaredInDocument    = false;
    bool                     detectedHeuristically = false;
};

struct LanguageTag { std::string bcp47; bool Empty() const { return bcp47.empty(); } };

enum class Script {
    Unknown, Latin, Hangul, Han, Hiragana, Katakana, Arabic, Cyrillic,
    Greek, Hebrew, Thai, Devanagari, Bengali, Tamil, Georgian, Armenian,
    Ethiopic, Symbol, Mixed
};

enum class FontCharset {
    Unknown, ANSI, Default, Symbol, Mac, ShiftJIS, Hangeul, GB2312, ChineseBig5,
    Greek, Turkish, Hebrew, Arabic, Baltic, Russian, Thai, EastEurope, OEM
};

enum class LineBreakMode { Unknown, Auto, Normal, KeepAll, BreakAll };
enum class WordBreakMode { Unknown, Normal, BreakAll, KeepAll };

struct TextEnvironment {
    std::optional<LanguageTag>  defaultLanguage;
    std::optional<Script>       defaultScript;
    std::optional<FontCharset>  defaultCharset;
    std::optional<EncodingInfo> sourceEncoding;
    std::optional<EncodingInfo> preferredSaveEncoding;
    std::optional<LineBreakMode> lineBreakMode;
    std::optional<WordBreakMode> wordBreakMode;
};

enum class Unit { Unknown, Px, Pt, Mm, Cm, Inch, Percent, Em, Rem };

struct Length {
    double value = 0.0;
    Unit   unit  = Unit::Unknown;
    static Length Px(double v)      { return {v, Unit::Px};      }
    static Length Pt(double v)      { return {v, Unit::Pt};      }
    static Length Mm(double v)      { return {v, Unit::Mm};      }
    static Length Cm(double v)      { return {v, Unit::Cm};      }
    static Length Inch(double v)    { return {v, Unit::Inch};    }
    static Length Percent(double v) { return {v, Unit::Percent}; }
    static Length Em(double v)      { return {v, Unit::Em};      }
    static Length Rem(double v)     { return {v, Unit::Rem};     }
};

struct Color {
    uint8_t r = 0, g = 0, b = 0, a = 255;
    static Color Make(uint8_t r, uint8_t g, uint8_t b, uint8_t a = 255) { return {r,g,b,a}; }
};

enum class WritingDirection { LTR, RTL, Vertical };
enum class Alignment { Unknown, Left, Center, Right, Justify };
enum class UnderlineStyle { None, Single, Double, Dotted, Dashed, Wavy };
enum class BorderStyle { None, Solid, Dashed, Dotted, Double };
enum class BreakType { LineBreak, SoftBreak, PageBreak, ColumnBreak, SectionBreak };
enum class ListType { None, Bullet, Numbered, RomanUpper, RomanLower, AlphaUpper, AlphaLower };
enum class TableLayoutMode { Auto, Fixed };
enum class ImageType { Unknown, Raster, Vector };
enum class NoteType { Footnote, Endnote };
enum class RawFragmentKind { Unknown, Html, Markdown, Xml, DocxXml, HwpxXml, PlainText };

struct SourceSpan { std::optional<size_t> start; std::optional<size_t> end; };
struct SourceInfo {
    FileFormat  sourceFormat = FileFormat::Unknown;
    std::string sourcePath;
    std::string originalNodeName;
    SourceSpan  span;
    std::optional<EncodingInfo> encoding;
};

struct ExtensionValue {
    using Value = std::variant<std::nullptr_t, bool, int64_t, double, std::string, BinaryBlob>;
    Value value;
};
struct ExtensionBag { std::unordered_map<std::string, ExtensionValue> values; };

struct NodeBase {
    NodeId       id = 0;
    SourceInfo   source;
    MetadataMap  metadata;
    ExtensionBag extensions;
    virtual ~NodeBase() = default;
};

struct Margin  { std::optional<Length> left, right, top, bottom; };
struct Padding { std::optional<Length> left, right, top, bottom; };
struct BorderEdge { BorderStyle style = BorderStyle::None; std::optional<Length> width; std::optional<Color> color; };
struct Border { BorderEdge left, right, top, bottom; };

struct TextStyle {
    std::optional<std::string>      fontFamily;
    std::optional<Length>           fontSize;
    std::optional<Color>            color;
    std::optional<Color>            backgroundColor;
    std::optional<bool>             bold;
    std::optional<bool>             italic;
    std::optional<bool>             strike;
    std::optional<bool>             doubleStrike;
    std::optional<bool>             outline;
    std::optional<bool>             shadow;
    std::optional<bool>             smallCaps;
    std::optional<bool>             allCaps;
    std::optional<UnderlineStyle>   underline;
    std::optional<Length>           characterSpacing;
    std::optional<LanguageTag>      language;
    std::optional<Script>           script;
    std::optional<FontCharset>      charset;
};

struct ParagraphStyle {
    std::optional<Alignment>       alignment;
    std::optional<Length>          indentLeft;
    std::optional<Length>          indentRight;
    std::optional<Length>          indentFirstLine;
    std::optional<Length>          spacingBefore;
    std::optional<Length>          spacingAfter;
    std::optional<Length>          lineSpacing;
    std::optional<WritingDirection> direction;
    std::optional<bool>            keepWithNext;
    std::optional<bool>            keepLinesTogether;
    std::optional<bool>            pageBreakBefore;
    std::optional<LanguageTag>     language;
    std::optional<LineBreakMode>   lineBreakMode;
    std::optional<WordBreakMode>   wordBreakMode;
};

struct BoxStyle { Margin margin; Padding padding; Border border; std::optional<Color> backgroundColor; std::optional<Length> width, height; };

struct StyleDefinition {
    StyleId id; std::string name;
    std::optional<StyleId> basedOn, nextStyle;
    TextStyle text; ParagraphStyle paragraph; BoxStyle box;
    MetadataMap metadata; ExtensionBag extensions;
};

struct Resource {
    ResourceId  id = 0;
    std::string name, mediaType, originalPath;
    BinaryBlob  data;
    MetadataMap metadata; ExtensionBag extensions;
};

struct Bookmark { std::string name; NodeId targetNodeId = 0; };

struct Comment {
    std::string commentId, author, initials, dateTime;
    std::vector<std::string> paragraphs;
    MetadataMap metadata; ExtensionBag extensions;
};

struct Note {
    std::string noteId;
    NoteType    type = NoteType::Footnote;
    MetadataMap metadata; ExtensionBag extensions;
};

// Forward declarations
struct Inline; struct Block;
struct TableCell; struct TableRow;
struct HeaderFooter; struct Section;

using InlinePtr      = std::shared_ptr<Inline>;
using BlockPtr       = std::shared_ptr<Block>;
using TableCellPtr   = std::shared_ptr<TableCell>;
using TableRowPtr    = std::shared_ptr<TableRow>;
using HeaderFooterPtr = std::shared_ptr<HeaderFooter>;
using SectionPtr     = std::shared_ptr<Section>;

// ---- Inline nodes ----
struct Text : NodeBase { std::u32string text; std::optional<StyleId> styleRef; TextStyle directStyle; std::optional<EncodingInfo> originalEncoding; };
struct Tab  : NodeBase {};
struct Break : NodeBase { BreakType type = BreakType::LineBreak; };
struct InlineCode : NodeBase { std::u32string text; std::optional<StyleId> styleRef; TextStyle directStyle; std::optional<EncodingInfo> originalEncoding; };
struct Hyperlink : NodeBase { std::string href, title; };
struct Image : NodeBase { std::optional<ResourceId> resourceId; std::string altText, title; std::optional<Length> width, height; ImageType imageType = ImageType::Unknown; };
struct Field : NodeBase { std::string fieldType, instruction, resultText; };
struct NoteRef : NodeBase { std::string noteId; NoteType type = NoteType::Footnote; };
struct CommentRangeStart : NodeBase { std::string commentId; };
struct CommentRangeEnd   : NodeBase { std::string commentId; };
struct BookmarkStart : NodeBase { std::string name; };
struct BookmarkEnd   : NodeBase { std::string name; };
struct RawInline : NodeBase { RawFragmentKind kind = RawFragmentKind::Unknown; std::string data; std::optional<EncodingInfo> encoding; };

struct InlineContainer : NodeBase {
    std::vector<InlinePtr> children;
    std::optional<StyleId> styleRef;
    TextStyle directStyle;
};
struct Span     : InlineContainer {};
struct Emphasis : InlineContainer {};
struct Strong   : InlineContainer {};
struct Underline: InlineContainer {};

struct Inline {
    using Variant = std::variant<
        Text, Tab, Break, InlineCode, Hyperlink, Image, Field,
        NoteRef, CommentRangeStart, CommentRangeEnd, BookmarkStart, BookmarkEnd,
        RawInline, Span, Emphasis, Strong, Underline>;
    Variant value;
    Inline() = default;
    template<typename T> Inline(T v) : value(std::move(v)) {}
};

// ---- Block nodes ----
struct Paragraph : NodeBase {
    std::vector<InlinePtr> inlines;
    std::optional<StyleId> styleRef;
    ParagraphStyle         directStyle;
    std::optional<ListType>     listType;
    std::optional<int32_t>      listLevel;
    std::optional<int32_t>      listStart;
    std::string                 listLabel;
};

struct Heading : NodeBase {
    int level = 1;
    std::vector<InlinePtr> inlines;
    std::optional<StyleId> styleRef;
    ParagraphStyle         directStyle;
};

struct CodeBlock : NodeBase {
    std::string code, language;
    std::optional<StyleId> styleRef;
    BoxStyle directStyle;
    std::optional<EncodingInfo> originalEncoding;
};

struct RawBlock : NodeBase { RawFragmentKind kind = RawFragmentKind::Unknown; std::string data; std::optional<EncodingInfo> encoding; };
struct HorizontalRule : NodeBase {};

struct TableCell : NodeBase { std::vector<BlockPtr> blocks; std::optional<int32_t> rowSpan, colSpan; BoxStyle directStyle; };
struct TableRow  : NodeBase { std::vector<TableCellPtr> cells; };
struct Table     : NodeBase { std::vector<TableRowPtr> rows; TableLayoutMode layoutMode = TableLayoutMode::Auto; std::optional<StyleId> styleRef; BoxStyle directStyle; };
struct BlockQuote: NodeBase { std::vector<BlockPtr> blocks; std::optional<StyleId> styleRef; BoxStyle directStyle; };
struct ListItem  : NodeBase { std::vector<BlockPtr> blocks; };
struct ListBlock : NodeBase { ListType type = ListType::Bullet; int32_t level = 0, start = 1; std::vector<std::shared_ptr<ListItem>> items; };
struct SectionBreakBlock : NodeBase {};
struct CustomBlock : NodeBase { std::string name; std::vector<BlockPtr> blocks; };

struct Block {
    using Variant = std::variant<
        Paragraph, Heading, CodeBlock, RawBlock, HorizontalRule,
        Table, BlockQuote, ListBlock, SectionBreakBlock, CustomBlock>;
    Variant value;
    Block() = default;
    template<typename T> Block(T v) : value(std::move(v)) {}
};

// ---- Page / Section ----
struct PageSize    { std::optional<Length> width, height; };
struct PageMargins { std::optional<Length> left, right, top, bottom, header, footer, gutter; };
struct PageSettings { PageSize size; PageMargins margins; std::optional<bool> landscape; };

struct HeaderFooter : NodeBase { std::vector<BlockPtr> blocks; };

struct Section : NodeBase {
    PageSettings pageSettings;
    HeaderFooterPtr headerDefault, headerFirst, headerEven;
    HeaderFooterPtr footerDefault, footerFirst, footerEven;
    std::vector<BlockPtr> blocks;
};

// ---- Document root ----
struct DocumentProperties {
    std::string title, subject, author, creator, keywords, description;
    std::string language, createdAt, modifiedAt;
};

struct Document {
    NodeId              id = 0;
    FileFormat          originalFormat = FileFormat::Unknown;
    DocumentProperties  properties;
    MetadataMap         metadata;
    ExtensionBag        extensions;
    TextEnvironment     textEnvironment;
    std::vector<StyleDefinition> styles;
    std::vector<Resource>        resources;
    std::vector<Comment>         comments;
    std::vector<Note>            notes;
    std::vector<Bookmark>        bookmarks;
    std::vector<SectionPtr>      sections;
};

// Returns pointer to the resource with matching id, or nullptr if absent.
// O(N) on resources; callers hitting this in a hot path should cache.
inline const Resource* FindResource(const Document& doc, ResourceId id) {
    for (const auto& r : doc.resources)
        if (r.id == id) return &r;
    return nullptr;
}

// ---- Factory helpers ----
class NodeIdGenerator { NodeId current_ = 0; public: NodeId Next() { return ++current_; } };

template<typename T, typename... Args>
std::shared_ptr<T> MakeNode(NodeIdGenerator& gen, Args&&... args) {
    auto node = std::make_shared<T>(std::forward<Args>(args)...);
    node->id = gen.Next();
    return node;
}

template<typename T, typename... Args>
InlinePtr MakeInline(NodeIdGenerator& gen, Args&&... args) {
    auto inl = std::make_shared<Inline>(T{std::forward<Args>(args)...});
    std::get<T>(inl->value).id = gen.Next();
    return inl;
}

template<typename T, typename... Args>
BlockPtr MakeBlock(NodeIdGenerator& gen, Args&&... args) {
    auto blk = std::make_shared<Block>(T{std::forward<Args>(args)...});
    std::get<T>(blk->value).id = gen.Next();
    return blk;
}

inline bool HasUnicodeEncoding(TextEncoding enc) {
    switch (enc) { case TextEncoding::UTF8: case TextEncoding::UTF8_BOM:
        case TextEncoding::UTF16LE: case TextEncoding::UTF16BE:
        case TextEncoding::UTF32LE: case TextEncoding::UTF32BE: return true;
        default: return false; }
}

} // namespace cdm
