#include "CdmNormalizer.hpp"
#include <vector>

namespace cdm {

// ---- Style equality helpers ------------------------------------------------

static bool ColorsEqual(const std::optional<Color>& a, const std::optional<Color>& b) {
    if (a.has_value() != b.has_value()) return false;
    if (!a) return true;
    return a->r == b->r && a->g == b->g && a->b == b->b && a->a == b->a;
}

static bool LengthsEqual(const std::optional<Length>& a, const std::optional<Length>& b) {
    if (a.has_value() != b.has_value()) return false;
    if (!a) return true;
    return a->value == b->value && a->unit == b->unit;
}

static bool LangsEqual(const std::optional<LanguageTag>& a, const std::optional<LanguageTag>& b) {
    if (a.has_value() != b.has_value()) return false;
    if (!a) return true;
    return a->bcp47 == b->bcp47;
}

static bool StylesEqual(const TextStyle& a, const TextStyle& b) {
    return a.fontFamily       == b.fontFamily       &&
           LengthsEqual(a.fontSize,          b.fontSize)          &&
           ColorsEqual(a.color,              b.color)              &&
           ColorsEqual(a.backgroundColor,    b.backgroundColor)    &&
           a.bold             == b.bold             &&
           a.italic           == b.italic           &&
           a.strike           == b.strike           &&
           a.doubleStrike     == b.doubleStrike     &&
           a.outline          == b.outline          &&
           a.shadow           == b.shadow           &&
           a.smallCaps        == b.smallCaps        &&
           a.allCaps          == b.allCaps          &&
           a.underline        == b.underline        &&
           LengthsEqual(a.characterSpacing, b.characterSpacing)    &&
           LangsEqual(a.language,           b.language)            &&
           a.script           == b.script           &&
           a.charset          == b.charset;
}

// ---- Traversal (forward declarations for mutual recursion) -----------------

static void NormalizeInlines(std::vector<InlinePtr>& inlines);
static void NormalizeBlocks(std::vector<BlockPtr>& blocks);

// ---- Inline merge ----------------------------------------------------------

static void MergeAdjacentText(std::vector<InlinePtr>& inlines) {
    if (inlines.size() < 2) return;
    std::vector<InlinePtr> out;
    out.reserve(inlines.size());
    for (auto& inl : inlines) {
        if (!out.empty()) {
            auto* prev = std::get_if<Text>(&out.back()->value);
            auto* cur  = std::get_if<Text>(&inl->value);
            if (prev && cur &&
                prev->styleRef   == cur->styleRef &&
                StylesEqual(prev->directStyle, cur->directStyle)) {
                prev->text += cur->text;
                continue;
            }
        }
        out.push_back(std::move(inl));
    }
    inlines = std::move(out);
}

static void NormalizeInlines(std::vector<InlinePtr>& inlines) {
    // Recurse into inline containers first (depth-first)
    for (auto& inl : inlines) {
        std::visit([](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_base_of_v<InlineContainer, T>)
                NormalizeInlines(v.children);
        }, inl->value);
    }
    MergeAdjacentText(inlines);
}

// ---- Block traversal -------------------------------------------------------

static void NormalizeBlocks(std::vector<BlockPtr>& blocks) {
    for (auto& blk : blocks) {
        std::visit([](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, Paragraph>) {
                NormalizeInlines(v.inlines);
            } else if constexpr (std::is_same_v<T, Heading>) {
                NormalizeInlines(v.inlines);
            } else if constexpr (std::is_same_v<T, BlockQuote>) {
                NormalizeBlocks(v.blocks);
            } else if constexpr (std::is_same_v<T, ListBlock>) {
                for (auto& item : v.items)
                    NormalizeBlocks(item->blocks);
            } else if constexpr (std::is_same_v<T, Table>) {
                for (auto& row : v.rows)
                    for (auto& cell : row->cells)
                        NormalizeBlocks(cell->blocks);
            } else if constexpr (std::is_same_v<T, CustomBlock>) {
                NormalizeBlocks(v.blocks);
            }
        }, blk->value);
    }
}

// ---- Public entry point ----------------------------------------------------

void Normalize(Document& doc) {
    for (auto& sec : doc.sections)
        NormalizeBlocks(sec->blocks);
}

} // namespace cdm
