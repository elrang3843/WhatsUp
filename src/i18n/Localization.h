#pragma once
#include <windows.h>
#include <string>
#include <unordered_map>

// ---- String ID enumeration ----
enum class SID {
    // App
    APP_TITLE,
    APP_VERSION,
    APP_COPYRIGHT,
    UNTITLED,

    // Menu: File
    MENU_FILE,
    MENU_FILE_NEW,
    MENU_FILE_OPEN,
    MENU_FILE_SAVE,
    MENU_FILE_SAVEAS,
    MENU_FILE_PRINT,
    MENU_FILE_CLOSE,
    MENU_FILE_EXIT,

    // Menu: Edit
    MENU_EDIT,
    MENU_EDIT_UNDO,
    MENU_EDIT_REDO,
    MENU_EDIT_CUT,
    MENU_EDIT_COPY,
    MENU_EDIT_PASTE,
    MENU_EDIT_FIND,
    MENU_EDIT_REPLACE,
    MENU_EDIT_DOCINFO,
    MENU_EDIT_SELECTALL,

    // Menu: Input (Insert)
    MENU_INPUT,
    MENU_INPUT_TEXTBOX,
    MENU_INPUT_CHARMAP,
    MENU_INPUT_SHAPES,
    MENU_INPUT_PICTURE,
    MENU_INPUT_EMOJI,
    MENU_INPUT_TABLE,
    MENU_INPUT_CHART,
    MENU_INPUT_FORMULA,
    MENU_INPUT_DATETIME,
    MENU_INPUT_COMMENT,
    MENU_INPUT_CAPTION,
    MENU_INPUT_HEADNOTE,
    MENU_INPUT_HYPERLINK,
    MENU_INPUT_BOOKMARK,

    // Menu: Format
    MENU_FORMAT,
    MENU_FORMAT_CHAR,
    MENU_FORMAT_PARA,
    MENU_FORMAT_COLUMNS,
    MENU_FORMAT_OBJECT,
    MENU_FORMAT_BORDER,
    MENU_FORMAT_STYLE,

    // Menu: Page
    MENU_PAGE,
    MENU_PAGE_SETUP,
    MENU_PAGE_HEADER,
    MENU_PAGE_FOOTER,
    MENU_PAGE_NUMBER,
    MENU_PAGE_SPLIT,
    MENU_PAGE_MERGE,

    // Menu: Settings
    MENU_SETTINGS,
    MENU_SETTINGS_GRID,
    MENU_SETTINGS_USER,
    MENU_SETTINGS_LANGUAGE,
    MENU_SETTINGS_THEME,

    // Menu: Help
    MENU_HELP,
    MENU_HELP_MANUAL,
    MENU_HELP_LICENSE,
    MENU_HELP_ABOUT,

    // Dialogs / UI text
    DLG_OK,
    DLG_CANCEL,
    DLG_YES,
    DLG_NO,
    DLG_CLOSE,

    MSG_UNSAVED_CHANGES,
    MSG_UNSAVED_TITLE,
    MSG_FILE_NOT_FOUND,
    MSG_OPEN_FAILED,
    MSG_SAVE_FAILED,
    MSG_PRINT_FAILED,
    MSG_FORMAT_UNSUPPORTED,
    MSG_HWP_READONLY,
    MSG_DOC_READONLY,
    MSG_SPELL_UNAVAILABLE,

    // Status bar
    STATUS_READY,
    STATUS_MODIFIED,
    STATUS_WORDS,
    STATUS_CHARS,
    STATUS_PAGE,
    STATUS_LINE,
    STATUS_COL,

    // File filter strings
    FILTER_ALL_DOCS,
    FILTER_TXT,
    FILTER_HTML,
    FILTER_MD,
    FILTER_DOCX,
    FILTER_HWPX,
    FILTER_HWP,
    FILTER_DOC,
    FILTER_ALL_FILES,

    // Tab labels
    TAB_GENERAL,
    TAB_USER,
    TAB_SPELL,
    TAB_GRID,
    TAB_PROPERTIES,
    TAB_STATISTICS,
    TAB_WATERMARK,

    // Toolbar tooltips
    TT_NEW,
    TT_OPEN,
    TT_SAVE,
    TT_PRINT,
    TT_UNDO,
    TT_REDO,
    TT_CUT,
    TT_COPY,
    TT_PASTE,
    TT_BOLD,
    TT_ITALIC,
    TT_UNDERLINE,
    TT_STRIKEOUT,
    TT_ALIGN_LEFT,
    TT_ALIGN_CENTER,
    TT_ALIGN_RIGHT,
    TT_ALIGN_JUSTIFY,
    TT_INDENT,
    TT_OUTDENT,
    TT_FONT_COLOR,
    TT_BGCOLOR,
    TT_FIND,
    TT_REPLACE,

    _COUNT
};

// ---- Language codes ----
enum class Language { Korean, English };

class Localization {
public:
    static void        SetLanguage(Language lang);
    static Language    GetLanguage();
    static const wchar_t* Get(SID id);

    // File open/save filter string (double-null terminated)
    static std::wstring GetFileFilter();

private:
    static Language                                         s_lang;
    static const wchar_t* const                             s_ko[];
    static const wchar_t* const                             s_en[];
};
