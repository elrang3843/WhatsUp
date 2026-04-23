#pragma once

// Icons
#define IDI_APPICON         1
#define IDI_APPICON_SMALL   2

// Main menu resource
#define IDR_MAINMENU        100
#define IDR_ACCELERATORS    101

// ---- File menu ----
#define ID_FILE_NEW         201
#define ID_FILE_OPEN        202
#define ID_FILE_SAVE        203
#define ID_FILE_SAVEAS      204
#define ID_FILE_PRINT       205
#define ID_FILE_CLOSE       206
#define ID_FILE_EXIT        207

// ---- Edit menu ----
#define ID_EDIT_UNDO        301
#define ID_EDIT_REDO        302
#define ID_EDIT_CUT         303
#define ID_EDIT_COPY        304
#define ID_EDIT_PASTE       305
#define ID_EDIT_FIND        306
#define ID_EDIT_REPLACE     307
#define ID_EDIT_DOCINFO     308
#define ID_EDIT_SELECTALL   309

// ---- Input (Insert) menu ----
#define ID_INSERT_TEXTBOX   401
#define ID_INSERT_CHARMAP   402
#define ID_INSERT_SHAPES    403
#define ID_INSERT_PICTURE   404
#define ID_INSERT_EMOJI     405
#define ID_INSERT_TABLE     406
#define ID_INSERT_CHART     407
#define ID_INSERT_FORMULA   408
#define ID_INSERT_DATETIME  409
#define ID_INSERT_COMMENT   410
#define ID_INSERT_CAPTION   411
#define ID_INSERT_HEADNOTE  412
#define ID_INSERT_HYPERLINK 413
#define ID_INSERT_BOOKMARK  414

// ---- Format menu ----
#define ID_FORMAT_CHAR      501
#define ID_FORMAT_PARA      502
#define ID_FORMAT_COLUMNS   503
#define ID_FORMAT_OBJECT    504
#define ID_FORMAT_BORDER    505
#define ID_FORMAT_STYLE     506

// Format toolbar commands
#define ID_FORMAT_BOLD          510
#define ID_FORMAT_ITALIC        511
#define ID_FORMAT_UNDERLINE     512
#define ID_FORMAT_STRIKEOUT     513
#define ID_FORMAT_SUPERSCRIPT   514
#define ID_FORMAT_SUBSCRIPT     515
#define ID_FORMAT_FONTCOLOR     516
#define ID_FORMAT_BGCOLOR       517
#define ID_FORMAT_OUTLINE       518

#define ID_PARA_ALIGN_LEFT      520
#define ID_PARA_ALIGN_CENTER    521
#define ID_PARA_ALIGN_RIGHT     522
#define ID_PARA_ALIGN_JUSTIFY   523
#define ID_PARA_INDENT          524
#define ID_PARA_OUTDENT         525
#define ID_PARA_NUMBERING       526
#define ID_PARA_BULLET          527

// ---- Page menu ----
#define ID_PAGE_SETUP           601
#define ID_PAGE_HEADER          602
#define ID_PAGE_FOOTER          603
#define ID_PAGE_NUMBER          604
#define ID_PAGE_SPLIT           605
#define ID_PAGE_MERGE           606

// ---- Settings menu ----
#define ID_SETTINGS_GRID        701
#define ID_SETTINGS_USER        702
#define ID_SETTINGS_LANGUAGE    703
#define ID_SETTINGS_THEME       704

// ---- Help menu ----
#define ID_HELP_MANUAL          801
#define ID_HELP_LICENSE         802
#define ID_HELP_ABOUT           803

// ---- Controls ----
#define IDC_EDITOR              1001
#define IDC_STATUSBAR           1002
#define IDC_TOOLBAR_STANDARD    1003
#define IDC_TOOLBAR_FORMAT      1004
#define IDC_REBAR               1005
#define IDC_COMBO_FONT          1006
#define IDC_COMBO_SIZE          1007
#define IDC_RULER               1008

// ---- Dialogs ----
#define IDD_ABOUT               2001
#define IDD_SETTINGS            2002
#define IDD_DOC_INFO            2003
#define IDD_INSERT_TABLE        2004
#define IDD_INSERT_HYPERLINK    2005
#define IDD_INSERT_DATETIME     2006
#define IDD_INSERT_COMMENT      2007
#define IDD_PAGE_NUMBER         2008
#define IDD_COLUMNS             2009
#define IDD_BORDER_SHADING      2010
#define IDD_STYLE               2011
#define IDD_INSERT_FORMULA      2012
#define IDD_HEADER_FOOTER       2013
#define IDD_CHARMAP             2014
#define IDD_INSERT_BOOKMARK     2015

// ---- Dialog Controls ----
// About dialog
#define IDC_ABOUT_LOGO          3001
#define IDC_ABOUT_VERSION       3002
#define IDC_ABOUT_COPYRIGHT     3003
#define IDC_ABOUT_URL           3004

// Settings dialog
#define IDC_SETTINGS_TABS       3010
#define IDC_LANG_COMBO          3011
#define IDC_THEME_COMBO         3012
#define IDC_USER_NAME           3013
#define IDC_USER_ORG            3014
#define IDC_SPELL_ENABLE        3015
#define IDC_SPELL_LANG          3016
#define IDC_AUTOCOMPLETE_ENABLE 3017
#define IDC_GRID_SHOW           3018
#define IDC_GRID_SNAP           3019
#define IDC_GRID_SIZE           3020

// Doc info dialog
#define IDC_DOCINFO_TABS        3030
#define IDC_DOC_TITLE           3031
#define IDC_DOC_AUTHOR          3032
#define IDC_DOC_SUBJECT         3033
#define IDC_DOC_KEYWORDS        3034
#define IDC_DOC_COMMENT         3035
#define IDC_WORD_COUNT          3036
#define IDC_CHAR_COUNT          3037
#define IDC_PAGE_COUNT          3038
#define IDC_WATERMARK_NONE      3039
#define IDC_WATERMARK_TEXT      3040
#define IDC_WATERMARK_IMAGE     3041
#define IDC_WATERMARK_CONTENT   3042
#define IDC_WATERMARK_OPACITY   3043
#define IDC_WATERMARK_ANGLE     3044
#define IDC_WATERMARK_COLOR     3045
#define IDC_WATERMARK_BROWSE    3046
#define IDC_WATERMARK_PREVIEW   3047

// Insert table
#define IDC_TABLE_ROWS          3060
#define IDC_TABLE_COLS          3061
#define IDC_TABLE_AUTOFIT       3062
#define IDC_TABLE_WIDTH_PCT     3063
#define IDC_TABLE_HAS_HEADER    3064

// Insert hyperlink
#define IDC_LINK_DISPLAY        3070
#define IDC_LINK_URL            3071
#define IDC_LINK_TOOLTIP        3072

// Insert date/time
#define IDC_DATETIME_LIST       3080
#define IDC_DATETIME_PREVIEW    3081
#define IDC_DATETIME_AUTO       3082

// Insert comment
#define IDC_COMMENT_TEXT        3090
#define IDC_COMMENT_AUTHOR      3091

// Page number
#define IDC_PAGENUM_POS         3100
#define IDC_PAGENUM_ALIGN       3101
#define IDC_PAGENUM_FORMAT      3102
#define IDC_PAGENUM_START       3103

// Columns
#define IDC_COL_COUNT           3110
#define IDC_COL_EQUAL           3111
#define IDC_COL_SEPARATOR       3112
#define IDC_COL_SPACING         3113

// Bookmark
#define IDC_BOOKMARK_LIST       3120
#define IDC_BOOKMARK_NAME       3121
#define IDC_BOOKMARK_ADD        3122
#define IDC_BOOKMARK_DELETE     3123
#define IDC_BOOKMARK_GOTO       3124

// Character map
#define IDC_CHARMAP_GRID        3130
#define IDC_CHARMAP_SUBSET      3131
#define IDC_CHARMAP_SEARCH      3132
#define IDC_CHARMAP_SELECTED    3133
#define IDC_CHARMAP_COPY        3134
#define IDC_CHARMAP_INSERT      3135

// String table (for version info)
#define IDS_APP_TITLE           4001
#define IDS_VERSION             4002
