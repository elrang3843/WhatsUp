#include "CdmNormalizer.hpp"
#include <cctype>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace cdm {

// ============================================================================
// Style equality (used by merge-adjacent-text pass)
// ============================================================================

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

// ============================================================================
// Style merge helpers
// ============================================================================
//
// Two merge modes:
//   * Overlay:  src overwrites dst where src has a value (leaf wins)
//   * Inherit:  src only fills empty slots on dst (dst / direct wins)
//
// The style-chain walk (root → leaf) uses Overlay so a derived style wins.
// The final merge of a resolved style into a node's directStyle uses Inherit
// so explicit direct formatting always wins over a referenced named style.
// ============================================================================

#define OVERLAY_FIELD(fld) if (src.fld) dst.fld = src.fld
#define INHERIT_FIELD(fld) if (!dst.fld && src.fld) dst.fld = src.fld

static void OverlayTextStyle(TextStyle& dst, const TextStyle& src) {
    OVERLAY_FIELD(fontFamily);       OVERLAY_FIELD(fontSize);
    OVERLAY_FIELD(color);            OVERLAY_FIELD(backgroundColor);
    OVERLAY_FIELD(bold);             OVERLAY_FIELD(italic);
    OVERLAY_FIELD(strike);           OVERLAY_FIELD(doubleStrike);
    OVERLAY_FIELD(outline);          OVERLAY_FIELD(shadow);
    OVERLAY_FIELD(smallCaps);        OVERLAY_FIELD(allCaps);
    OVERLAY_FIELD(underline);        OVERLAY_FIELD(characterSpacing);
    OVERLAY_FIELD(language);         OVERLAY_FIELD(script);
    OVERLAY_FIELD(charset);
}

static void InheritTextStyle(TextStyle& dst, const TextStyle& src) {
    INHERIT_FIELD(fontFamily);       INHERIT_FIELD(fontSize);
    INHERIT_FIELD(color);            INHERIT_FIELD(backgroundColor);
    INHERIT_FIELD(bold);             INHERIT_FIELD(italic);
    INHERIT_FIELD(strike);           INHERIT_FIELD(doubleStrike);
    INHERIT_FIELD(outline);          INHERIT_FIELD(shadow);
    INHERIT_FIELD(smallCaps);        INHERIT_FIELD(allCaps);
    INHERIT_FIELD(underline);        INHERIT_FIELD(characterSpacing);
    INHERIT_FIELD(language);         INHERIT_FIELD(script);
    INHERIT_FIELD(charset);
}

static void OverlayParaStyle(ParagraphStyle& dst, const ParagraphStyle& src) {
    OVERLAY_FIELD(alignment);        OVERLAY_FIELD(indentLeft);
    OVERLAY_FIELD(indentRight);      OVERLAY_FIELD(indentFirstLine);
    OVERLAY_FIELD(spacingBefore);    OVERLAY_FIELD(spacingAfter);
    OVERLAY_FIELD(lineSpacing);      OVERLAY_FIELD(direction);
    OVERLAY_FIELD(keepWithNext);     OVERLAY_FIELD(keepLinesTogether);
    OVERLAY_FIELD(pageBreakBefore);  OVERLAY_FIELD(language);
    OVERLAY_FIELD(lineBreakMode);    OVERLAY_FIELD(wordBreakMode);
}

static void InheritParaStyle(ParagraphStyle& dst, const ParagraphStyle& src) {
    INHERIT_FIELD(alignment);        INHERIT_FIELD(indentLeft);
    INHERIT_FIELD(indentRight);      INHERIT_FIELD(indentFirstLine);
    INHERIT_FIELD(spacingBefore);    INHERIT_FIELD(spacingAfter);
    INHERIT_FIELD(lineSpacing);      INHERIT_FIELD(direction);
    INHERIT_FIELD(keepWithNext);     INHERIT_FIELD(keepLinesTogether);
    INHERIT_FIELD(pageBreakBefore);  INHERIT_FIELD(language);
    INHERIT_FIELD(lineBreakMode);    INHERIT_FIELD(wordBreakMode);
}

#undef OVERLAY_FIELD
#undef INHERIT_FIELD

// ============================================================================
// Style resolver — walks styleRef + basedOn chain and produces a flat style
// ============================================================================

using StyleMap = std::unordered_map<std::string, const StyleDefinition*>;

struct ResolvedStyle {
    TextStyle       text;
    ParagraphStyle  paragraph;
    std::string     resolvedName;   // name of the leaf style (if any)
};

static StyleMap BuildStyleMap(const Document& doc) {
    StyleMap m;
    m.reserve(doc.styles.size());
    for (const auto& s : doc.styles) m[s.id] = &s;
    return m;
}

// Resolve a styleRef by walking basedOn ancestors. Returns the fully merged
// text + paragraph style with leaf-winning semantics. Cycles are guarded.
static ResolvedStyle Resolve(const StyleId& id, const StyleMap& map) {
    ResolvedStyle r;

    // Collect chain leaf → root, then reverse to root → leaf for overlay.
    std::vector<const StyleDefinition*> chain;
    std::unordered_set<std::string> seen;
    std::string cur = id;
    while (!cur.empty() && !seen.count(cur)) {
        seen.insert(cur);
        auto it = map.find(cur);
        if (it == map.end()) break;
        chain.push_back(it->second);
        cur = it->second->basedOn ? *it->second->basedOn : std::string{};
    }
    // Walk root → leaf so that leaf overrides base.
    for (auto it = chain.rbegin(); it != chain.rend(); ++it) {
        OverlayTextStyle(r.text, (*it)->text);
        OverlayParaStyle(r.paragraph, (*it)->paragraph);
    }
    if (!chain.empty()) r.resolvedName = chain.front()->name;
    return r;
}

// ============================================================================
// Stage 1: Resolve style refs → merge into directStyle (direct wins)
// ============================================================================

static void ResolveInlines(std::vector<InlinePtr>& inlines,
                           const TextStyle&        inheritedText,
                           const StyleMap&         map);

static void ResolveInline(Inline& inl,
                          const TextStyle& inheritedText,
                          const StyleMap&  map)
{
    std::visit([&](auto& v) {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, Text> || std::is_same_v<T, InlineCode>) {
            if (v.styleRef) {
                auto rs = Resolve(*v.styleRef, map);
                InheritTextStyle(v.directStyle, rs.text);
            }
            // Fold in the paragraph's / ancestor's inherited text style.
            InheritTextStyle(v.directStyle, inheritedText);
        }
        else if constexpr (std::is_base_of_v<InlineContainer, T>) {
            TextStyle cascade = inheritedText;
            if (v.styleRef) {
                auto rs = Resolve(*v.styleRef, map);
                InheritTextStyle(v.directStyle, rs.text);
                // Build the cascade for children: container's (direct + resolved)
                OverlayTextStyle(cascade, rs.text);
            }
            OverlayTextStyle(cascade, v.directStyle);
            ResolveInlines(v.children, cascade, map);
        }
        // Tab, Break, Image, etc. have no styleRef/directStyle → nothing to do.
    }, inl.value);
}

static void ResolveInlines(std::vector<InlinePtr>& inlines,
                           const TextStyle&        inheritedText,
                           const StyleMap&         map)
{
    for (auto& inl : inlines) ResolveInline(*inl, inheritedText, map);
}

static void ResolveBlocks(std::vector<BlockPtr>& blocks, const StyleMap& map);

static void ResolveBlock(Block& blk, const StyleMap& map) {
    std::visit([&](auto& v) {
        using T = std::decay_t<decltype(v)>;

        if constexpr (std::is_same_v<T, Paragraph> || std::is_same_v<T, Heading>) {
            TextStyle inherited;
            if (v.styleRef) {
                auto rs = Resolve(*v.styleRef, map);
                InheritParaStyle(v.directStyle, rs.paragraph);
                inherited = rs.text;   // cascades into Text children
            }
            ResolveInlines(v.inlines, inherited, map);
        }
        else if constexpr (std::is_same_v<T, CodeBlock>) {
            // CodeBlock has no TextStyle directStyle to merge into; skip.
        }
        else if constexpr (std::is_same_v<T, BlockQuote>) {
            ResolveBlocks(v.blocks, map);
        }
        else if constexpr (std::is_same_v<T, ListBlock>) {
            for (auto& item : v.items) ResolveBlocks(item->blocks, map);
        }
        else if constexpr (std::is_same_v<T, Table>) {
            for (auto& row : v.rows)
                for (auto& cell : row->cells)
                    ResolveBlocks(cell->blocks, map);
        }
        else if constexpr (std::is_same_v<T, CustomBlock>) {
            ResolveBlocks(v.blocks, map);
        }
    }, blk.value);
}

static void ResolveBlocks(std::vector<BlockPtr>& blocks, const StyleMap& map) {
    for (auto& blk : blocks) ResolveBlock(*blk, map);
}

// ============================================================================
// Stage 2: Promote heading-styled paragraphs to Heading blocks
// ============================================================================

// Normalize a style name for matching: lowercase, strip spaces/underscores/hyphens.
static std::string NormName(const std::string& s) {
    std::string r;
    r.reserve(s.size());
    for (char c : s) {
        if (c == ' ' || c == '_' || c == '-') continue;
        r += static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    }
    return r;
}

// Returns 1-6 for heading styles, 0 otherwise.
// Matches "Heading1"…"Heading6" (case-insensitive, optional spacing).
// Also promotes common aliases: "Title"→1, "Subtitle"→2.
static int MatchHeadingLevel(const std::string& name) {
    std::string n = NormName(name);
    if (n == "title")    return 1;
    if (n == "subtitle") return 2;
    if (n.size() == 8 && n.substr(0, 7) == "heading") {
        char ch = n[7];
        if (ch >= '1' && ch <= '6') return ch - '0';
    }
    return 0;
}

static int HeadingLevelFromStyleRef(const std::optional<StyleId>& ref,
                                    const StyleMap&               map)
{
    if (!ref) return 0;
    int lvl = MatchHeadingLevel(*ref);
    if (lvl > 0) return lvl;
    auto it = map.find(*ref);
    if (it != map.end() && !it->second->name.empty())
        return MatchHeadingLevel(it->second->name);
    return 0;
}

static void PromoteHeadings(std::vector<BlockPtr>& blocks, const StyleMap& map) {
    for (auto& blk : blocks) {
        // Promote Paragraph → Heading in-place if styleRef matches.
        if (auto* p = std::get_if<Paragraph>(&blk->value)) {
            int lvl = HeadingLevelFromStyleRef(p->styleRef, map);
            if (lvl > 0) {
                Heading h;
                h.id          = p->id;
                h.source      = std::move(p->source);
                h.metadata    = std::move(p->metadata);
                h.extensions  = std::move(p->extensions);
                h.level       = lvl;
                h.inlines     = std::move(p->inlines);
                h.styleRef    = p->styleRef;
                h.directStyle = std::move(p->directStyle);
                blk->value    = std::move(h);
            }
        }

        // Recurse into container blocks (whichever variant we now hold).
        std::visit([&](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, BlockQuote>) PromoteHeadings(v.blocks, map);
            else if constexpr (std::is_same_v<T, ListBlock>)
                for (auto& item : v.items) PromoteHeadings(item->blocks, map);
            else if constexpr (std::is_same_v<T, Table>)
                for (auto& row : v.rows)
                    for (auto& cell : row->cells) PromoteHeadings(cell->blocks, map);
            else if constexpr (std::is_same_v<T, CustomBlock>) PromoteHeadings(v.blocks, map);
        }, blk->value);
    }
}

// ============================================================================
// Stage 3: Merge adjacent text inlines (existing)
// ============================================================================

static void NormalizeInlines(std::vector<InlinePtr>& inlines);
static void NormalizeBlocks(std::vector<BlockPtr>& blocks);

static void MergeAdjacentText(std::vector<InlinePtr>& inlines) {
    if (inlines.size() < 2) return;
    std::vector<InlinePtr> out;
    out.reserve(inlines.size());
    for (auto& inl : inlines) {
        if (!out.empty()) {
            auto* prev = std::get_if<Text>(&out.back()->value);
            auto* cur  = std::get_if<Text>(&inl->value);
            if (prev && cur &&
                prev->styleRef == cur->styleRef &&
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
    for (auto& inl : inlines) {
        std::visit([](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_base_of_v<InlineContainer, T>)
                NormalizeInlines(v.children);
        }, inl->value);
    }
    MergeAdjacentText(inlines);
}

static void NormalizeBlocks(std::vector<BlockPtr>& blocks) {
    for (auto& blk : blocks) {
        std::visit([](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, Paragraph> || std::is_same_v<T, Heading>) {
                NormalizeInlines(v.inlines);
            } else if constexpr (std::is_same_v<T, BlockQuote>) {
                NormalizeBlocks(v.blocks);
            } else if constexpr (std::is_same_v<T, ListBlock>) {
                for (auto& item : v.items) NormalizeBlocks(item->blocks);
            } else if constexpr (std::is_same_v<T, Table>) {
                for (auto& row : v.rows)
                    for (auto& cell : row->cells) NormalizeBlocks(cell->blocks);
            } else if constexpr (std::is_same_v<T, CustomBlock>) {
                NormalizeBlocks(v.blocks);
            }
        }, blk->value);
    }
}

// ============================================================================
// Stage 4: Collapse runs of empty paragraphs
// ============================================================================

static bool IsEmptyParagraph(const Block& blk) {
    auto* p = std::get_if<Paragraph>(&blk.value);
    if (!p) return false;
    if (p->inlines.empty()) return true;
    for (const auto& inl : p->inlines) {
        auto* t = std::get_if<Text>(&inl->value);
        if (!t) return false;             // any non-Text inline → not empty
        if (!t->text.empty()) return false;
    }
    return true;
}

static void CollapseEmptyParagraphs(std::vector<BlockPtr>& blocks) {
    std::vector<BlockPtr> out;
    out.reserve(blocks.size());
    bool prevEmpty = false;
    for (auto& b : blocks) {
        bool cur = IsEmptyParagraph(*b);
        if (cur && prevEmpty) continue;   // drop consecutive empties
        out.push_back(std::move(b));
        prevEmpty = cur;
    }
    blocks = std::move(out);

    // Recurse into container blocks.
    for (auto& b : blocks) {
        std::visit([](auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, BlockQuote>) CollapseEmptyParagraphs(v.blocks);
            else if constexpr (std::is_same_v<T, ListBlock>)
                for (auto& item : v.items) CollapseEmptyParagraphs(item->blocks);
            else if constexpr (std::is_same_v<T, Table>)
                for (auto& row : v.rows)
                    for (auto& cell : row->cells) CollapseEmptyParagraphs(cell->blocks);
            else if constexpr (std::is_same_v<T, CustomBlock>) CollapseEmptyParagraphs(v.blocks);
        }, b->value);
    }
}

// ============================================================================
// Public entry point
// ============================================================================

void Normalize(Document& doc) {
    StyleMap map = BuildStyleMap(doc);

    for (auto& sec : doc.sections) {
        ResolveBlocks(sec->blocks, map);          // 1. resolve styleRefs
        PromoteHeadings(sec->blocks, map);        // 2. Heading1-6 → Heading block
        NormalizeBlocks(sec->blocks);             // 3. merge adjacent text
        CollapseEmptyParagraphs(sec->blocks);     // 4. collapse blank paragraphs
    }
}

} // namespace cdm
