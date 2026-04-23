#include "MainWindow.h"
#include "Application.h"
#include "resource.h"
#include "i18n/Localization.h"
#include "formats/FormatManager.h"
#include "dialogs/AboutDialog.h"
#include "dialogs/SettingsDialog.h"
#include "dialogs/DocInfoDialog.h"
#include "dialogs/InsertTableDialog.h"
#include "dialogs/InsertHyperlinkDialog.h"
#include "dialogs/InsertDateTimeDialog.h"
#include <commdlg.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shellapi.h>
#include <richedit.h>
#include <cstdio>
#include <algorithm>

MainWindow* MainWindow::s_instance = nullptr;
UINT        MainWindow::s_findMsgId = 0;

// ─────────────────────────────────────────────
// Static factory
// ─────────────────────────────────────────────
HWND MainWindow::Create(HINSTANCE hInst) {
    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);
    wc.lpszClassName = kClassName;
    wc.hIcon         = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APPICON));
    wc.hIconSm       = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APPICON_SMALL));
    RegisterClassExW(&wc);

    s_findMsgId = RegisterWindowMessageW(FINDMSGSTRINGW);

    HWND hwnd = CreateWindowExW(
        WS_EX_ACCEPTFILES,
        kClassName,
        Localization::Get(StrID::APP_TITLE),
        WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
        CW_USEDEFAULT, CW_USEDEFAULT, 1100, 780,
        nullptr, nullptr, hInst, nullptr);
    return hwnd;
}

// ─────────────────────────────────────────────
// Window procedure
// ─────────────────────────────────────────────
LRESULT CALLBACK MainWindow::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    MainWindow* self = s_instance;

    // Modeless find/replace message
    if (msg == s_findMsgId && s_findMsgId != 0) {
        if (self) self->OnFindMsg(reinterpret_cast<FINDREPLACEW*>(lParam));
        return 0;
    }

    switch (msg) {
    case WM_CREATE: {
        auto* self2 = new MainWindow();
        s_instance  = self2;
        self2->m_hwnd = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self2));
        if (!self2->OnCreate(hwnd,
                reinterpret_cast<CREATESTRUCTW*>(lParam)->hInstance))
            return -1;
        return 0;
    }
    case WM_APP + 1:
        // Deferred editor init: RichEdit DLL's D2D thread must complete
        // before CreateWindowExW succeeds; posting this message lets
        // WM_CREATE finish first so the DLL is fully ready.
        if (self) {
            if (!self->m_editor->Create(hwnd, IDC_EDITOR)) {
                MessageBoxW(hwnd,
                    L"리치 에디트 컨트롤(RichEdit)을 초기화하지 못했습니다.\n"
                    L"msftedit.dll 또는 riched20.dll 로드 실패.\n\n"
                    L"시스템 파일이 손상되었거나 DirectX/D2D가 비활성화된 환경일 수 있습니다.",
                    L"WhatsUp 초기화 오류", MB_OK | MB_ICONERROR);
                DestroyWindow(hwnd);
            } else {
                self->m_ac->Create(hwnd, self->m_editor->GetHwnd());
                RECT rc{};
                GetClientRect(hwnd, &rc);
                self->OnSize(rc.right, rc.bottom);
                // Apply theme colors and refresh status/toolbar now that editor exists
                self->ApplyTheme();
                self->UpdateStatusBar();
                self->UpdateToolbarState();
                // Show the window now that everything is laid out. Doing it
                // here (not in Application::Init) means the very first paint
                // sees the correct rebar height and a live editor, so the
                // toolbar is immediately visible without a mouse-hover nudge.
                ShowWindow(hwnd, Application::Instance().GetNCmdShow());
                UpdateWindow(hwnd);
                SetFocus(self->m_editor->GetHwnd());
            }
        }
        return 0;
    case WM_SIZE:
        if (self) self->OnSize(LOWORD(lParam), HIWORD(lParam));
        return 0;
    case WM_COMMAND:
        if (self) self->OnCommand(LOWORD(wParam));
        return 0;
    case WM_NOTIFY:
        if (self) self->OnNotify(reinterpret_cast<NMHDR*>(lParam));
        return 0;
    case WM_CLOSE:
        if (self && !self->ConfirmClose()) return 0;
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY: {
        bool wasFullyCreated = self && self->m_fullyCreated;
        // Null s_instance BEFORE delete so any re-entrant WndProc calls
        // (e.g. EN_KILLFOCUS from the RichEdit's own WM_DESTROY) see nullptr
        // and skip processing rather than touching the half-destroyed object.
        s_instance = nullptr;
        delete self;
        if (wasFullyCreated)
            PostQuitMessage(0);
        return 0;
    }
    case WM_DROPFILES: {
        if (!self) break;
        HDROP hDrop = reinterpret_cast<HDROP>(wParam);
        wchar_t path[MAX_PATH];
        if (DragQueryFileW(hDrop, 0, path, MAX_PATH) > 0)
            self->OpenFile(path);
        DragFinish(hDrop);
        return 0;
    }
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ─────────────────────────────────────────────
// OnCreate
// ─────────────────────────────────────────────
bool MainWindow::OnCreate(HWND hwnd, HINSTANCE hInst) {
#define WU_LOG(s) OutputDebugStringW(L"[WhatsUp] " s L"\n")
    WU_LOG("OnCreate start");
    m_doc   = std::make_unique<Document>();   WU_LOG("Document OK");
    m_editor= std::make_unique<Editor>();      WU_LOG("Editor obj OK");
    m_spell = std::make_unique<SpellChecker>(); WU_LOG("SpellChecker OK");
    m_ac    = std::make_unique<AutoComplete>(); WU_LOG("AutoComplete OK");

    // Editor::Create() is deferred to WM_APP+1 so msftedit.dll's D2D
    // init thread has completed before CreateWindowExW is called.

    WU_LOG("BuildMenu...");       BuildMenu();
    WU_LOG("CreateToolbars...");  CreateToolbars();
    WU_LOG("CreateStatusBar..."); CreateStatusBar();
    WU_LOG("PopulateFontCombo..."); PopulateFontCombo();
    WU_LOG("PopulateSizeCombo..."); PopulateSizeCombo();

    auto& s = Application::Instance().Settings();
    m_spell->SetLanguage(s.spellLanguage); WU_LOG("SetLanguage OK");

    WU_LOG("ApplyTheme...");     ApplyTheme();
    WU_LOG("UpdateTitleBar..."); UpdateTitleBar();
    WU_LOG("UpdateStatusBar..."); UpdateStatusBar();

    m_fullyCreated = true;
    PostMessageW(hwnd, WM_APP + 1, 0, 0);
    WU_LOG("OnCreate done -> return true");
#undef WU_LOG
    return true;
}

// ─────────────────────────────────────────────
// BuildMenu — programmatic (Unicode-safe)
// ─────────────────────────────────────────────
void MainWindow::BuildMenu() {
    HMENU hBar  = CreateMenu();

    auto popup = [&](HMENU hParent, StrID nameId) {
        HMENU hPop = CreatePopupMenu();
        AppendMenuW(hParent, MF_POPUP, reinterpret_cast<UINT_PTR>(hPop),
                    Localization::Get(nameId));
        return hPop;
    };
    auto item = [](HMENU hPop, UINT id, StrID sid) {
        AppendMenuW(hPop, MF_STRING, id, Localization::Get(sid));
    };
    auto sep = [](HMENU hPop) {
        AppendMenuW(hPop, MF_SEPARATOR, 0, nullptr);
    };

    // 파일(File)
    HMENU hFile = popup(hBar, StrID::MENU_FILE);
    item(hFile, ID_FILE_NEW,    StrID::MENU_FILE_NEW);
    item(hFile, ID_FILE_OPEN,   StrID::MENU_FILE_OPEN);
    sep(hFile);
    item(hFile, ID_FILE_SAVE,   StrID::MENU_FILE_SAVE);
    item(hFile, ID_FILE_SAVEAS, StrID::MENU_FILE_SAVEAS);
    sep(hFile);
    item(hFile, ID_FILE_PRINT,  StrID::MENU_FILE_PRINT);
    sep(hFile);
    item(hFile, ID_FILE_CLOSE,  StrID::MENU_FILE_CLOSE);
    item(hFile, ID_FILE_EXIT,   StrID::MENU_FILE_EXIT);

    // 편집(Edit)
    HMENU hEdit = popup(hBar, StrID::MENU_EDIT);
    item(hEdit, ID_EDIT_UNDO,     StrID::MENU_EDIT_UNDO);
    item(hEdit, ID_EDIT_REDO,     StrID::MENU_EDIT_REDO);
    sep(hEdit);
    item(hEdit, ID_EDIT_CUT,      StrID::MENU_EDIT_CUT);
    item(hEdit, ID_EDIT_COPY,     StrID::MENU_EDIT_COPY);
    item(hEdit, ID_EDIT_PASTE,    StrID::MENU_EDIT_PASTE);
    sep(hEdit);
    item(hEdit, ID_EDIT_SELECTALL,StrID::MENU_EDIT_SELECTALL);
    sep(hEdit);
    item(hEdit, ID_EDIT_FIND,     StrID::MENU_EDIT_FIND);
    item(hEdit, ID_EDIT_REPLACE,  StrID::MENU_EDIT_REPLACE);
    sep(hEdit);
    item(hEdit, ID_EDIT_DOCINFO,  StrID::MENU_EDIT_DOCINFO);

    // 입력(Insert)
    HMENU hIns = popup(hBar, StrID::MENU_INPUT);
    item(hIns, ID_INSERT_TEXTBOX,  StrID::MENU_INPUT_TEXTBOX);
    item(hIns, ID_INSERT_CHARMAP,  StrID::MENU_INPUT_CHARMAP);
    item(hIns, ID_INSERT_SHAPES,   StrID::MENU_INPUT_SHAPES);
    item(hIns, ID_INSERT_PICTURE,  StrID::MENU_INPUT_PICTURE);
    item(hIns, ID_INSERT_EMOJI,    StrID::MENU_INPUT_EMOJI);
    sep(hIns);
    item(hIns, ID_INSERT_TABLE,    StrID::MENU_INPUT_TABLE);
    item(hIns, ID_INSERT_CHART,    StrID::MENU_INPUT_CHART);
    item(hIns, ID_INSERT_FORMULA,  StrID::MENU_INPUT_FORMULA);
    sep(hIns);
    item(hIns, ID_INSERT_DATETIME, StrID::MENU_INPUT_DATETIME);
    item(hIns, ID_INSERT_COMMENT,  StrID::MENU_INPUT_COMMENT);
    item(hIns, ID_INSERT_CAPTION,  StrID::MENU_INPUT_CAPTION);
    item(hIns, ID_INSERT_HEADNOTE, StrID::MENU_INPUT_HEADNOTE);
    sep(hIns);
    item(hIns, ID_INSERT_HYPERLINK,StrID::MENU_INPUT_HYPERLINK);
    item(hIns, ID_INSERT_BOOKMARK, StrID::MENU_INPUT_BOOKMARK);

    // 서식(Format)
    HMENU hFmt = popup(hBar, StrID::MENU_FORMAT);
    item(hFmt, ID_FORMAT_CHAR,    StrID::MENU_FORMAT_CHAR);
    item(hFmt, ID_FORMAT_PARA,    StrID::MENU_FORMAT_PARA);
    item(hFmt, ID_FORMAT_COLUMNS, StrID::MENU_FORMAT_COLUMNS);
    sep(hFmt);
    item(hFmt, ID_FORMAT_OBJECT,  StrID::MENU_FORMAT_OBJECT);
    item(hFmt, ID_FORMAT_BORDER,  StrID::MENU_FORMAT_BORDER);
    sep(hFmt);
    item(hFmt, ID_FORMAT_STYLE,   StrID::MENU_FORMAT_STYLE);

    // 쪽(Page)
    HMENU hPage = popup(hBar, StrID::MENU_PAGE);
    item(hPage, ID_PAGE_SETUP,  StrID::MENU_PAGE_SETUP);
    sep(hPage);
    item(hPage, ID_PAGE_HEADER, StrID::MENU_PAGE_HEADER);
    item(hPage, ID_PAGE_FOOTER, StrID::MENU_PAGE_FOOTER);
    item(hPage, ID_PAGE_NUMBER, StrID::MENU_PAGE_NUMBER);
    sep(hPage);
    item(hPage, ID_PAGE_SPLIT,  StrID::MENU_PAGE_SPLIT);
    item(hPage, ID_PAGE_MERGE,  StrID::MENU_PAGE_MERGE);

    // 설정(Settings)
    HMENU hSet = popup(hBar, StrID::MENU_SETTINGS);
    item(hSet, ID_SETTINGS_GRID,    StrID::MENU_SETTINGS_GRID);
    sep(hSet);
    item(hSet, ID_SETTINGS_USER,    StrID::MENU_SETTINGS_USER);
    item(hSet, ID_SETTINGS_LANGUAGE,StrID::MENU_SETTINGS_LANGUAGE);
    item(hSet, ID_SETTINGS_THEME,   StrID::MENU_SETTINGS_THEME);

    // 도움말(Help)
    HMENU hHelp = popup(hBar, StrID::MENU_HELP);
    item(hHelp, ID_HELP_MANUAL,  StrID::MENU_HELP_MANUAL);
    item(hHelp, ID_HELP_LICENSE, StrID::MENU_HELP_LICENSE);
    sep(hHelp);
    item(hHelp, ID_HELP_ABOUT,   StrID::MENU_HELP_ABOUT);

    SetMenu(m_hwnd, hBar);
}

// ─────────────────────────────────────────────
// CreateToolbars
// ─────────────────────────────────────────────
void MainWindow::CreateToolbars() {
    HINSTANCE hInst = Application::Instance().GetHInstance();

    // ── Rebar container ──────────────────────
    m_hwndRebar = CreateWindowExW(WS_EX_TOOLWINDOW,
        REBARCLASSNAMEW, nullptr,
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
        | RBS_VARHEIGHT | RBS_BANDBORDERS | CCS_NODIVIDER,
        0, 0, 0, 0, m_hwnd,
        reinterpret_cast<HMENU>(IDC_REBAR), hInst, nullptr);

    // ── Standard toolbar ─────────────────────
    m_hwndStdTb = CreateWindowExW(0, TOOLBARCLASSNAMEW, nullptr,
        WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS
        | CCS_NODIVIDER | CCS_NORESIZE | CCS_NOPARENTALIGN,
        0, 0, 0, 0, m_hwndRebar,
        reinterpret_cast<HMENU>(IDC_TOOLBAR_STANDARD), hInst, nullptr);

    SendMessageW(m_hwndStdTb, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
    SendMessageW(m_hwndStdTb, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DRAWDDARROWS);

    // Use standard shell image list (16x16)
    HIMAGELIST hImgStd = ImageList_Create(20, 20, ILC_COLOR32 | ILC_MASK, 16, 0);
    auto addSysIcon = [&](LPWSTR resId) {
        HICON hIco = reinterpret_cast<HICON>(
            LoadImageW(nullptr, resId, IMAGE_ICON, 20, 20, LR_SHARED));
        ImageList_AddIcon(hImgStd, hIco);
    };
    // Index order matches TBBUTTON iImage below
    addSysIcon(IDI_APPLICATION); // 0 placeholder — we'll use STD bitmaps instead

    // Use built-in STD bitmap
    TBADDBITMAP ab{ HINST_COMMCTRL, IDB_STD_SMALL_COLOR };
    SendMessageW(m_hwndStdTb, TB_ADDBITMAP, 0, reinterpret_cast<LPARAM>(&ab));

    TBBUTTON stdBtns[] = {
        {STD_FILENEW,   ID_FILE_NEW,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_FILEOPEN,  ID_FILE_OPEN,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_FILESAVE,  ID_FILE_SAVE,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_PRINT,     ID_FILE_PRINT, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {0, 0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {STD_UNDO,      ID_EDIT_UNDO,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_REDOW,     ID_EDIT_REDO,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {0, 0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {STD_CUT,       ID_EDIT_CUT,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_COPY,      ID_EDIT_COPY,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
        {STD_PASTE,     ID_EDIT_PASTE, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON, {}, 0, 0},
    };
    SendMessageW(m_hwndStdTb, TB_ADDBUTTONSW,
                 ARRAYSIZE(stdBtns), reinterpret_cast<LPARAM>(stdBtns));
    SendMessageW(m_hwndStdTb, TB_AUTOSIZE, 0, 0);

    // ── Format toolbar ────────────────────────
    m_hwndFmtTb = CreateWindowExW(0, TOOLBARCLASSNAMEW, nullptr,
        WS_CHILD | WS_VISIBLE | TBSTYLE_FLAT | TBSTYLE_TOOLTIPS | TBSTYLE_LIST
        | CCS_NODIVIDER | CCS_NORESIZE | CCS_NOPARENTALIGN,
        0, 0, 0, 0, m_hwndRebar,
        reinterpret_cast<HMENU>(IDC_TOOLBAR_FORMAT), hInst, nullptr);
    SendMessageW(m_hwndFmtTb, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

    // Font name combo embedded in toolbar
    m_hwndFontCombo = CreateWindowExW(0, L"COMBOBOX", nullptr,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_HASSTRINGS | WS_VSCROLL,
        0, 1, 200, 200, m_hwndFmtTb,
        reinterpret_cast<HMENU>(IDC_COMBO_FONT), hInst, nullptr);

    // Size combo
    m_hwndSizeCombo = CreateWindowExW(0, L"COMBOBOX", nullptr,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | CBS_HASSTRINGS | WS_VSCROLL,
        208, 1, 54, 200, m_hwndFmtTb,
        reinterpret_cast<HMENU>(IDC_COMBO_SIZE), hInst, nullptr);

    // Use text-drawing bitmaps for B/I/U etc.
    // ILC_COLOR24: avoids 32-bpp alpha-channel issues where ImageList_AddMasked
    // would see alpha=0 on every pixel of a plain DDB and treat them all as
    // transparent, rendering nothing (appears as solid black squares).
    HIMAGELIST hFmtImg = ImageList_Create(20, 20, ILC_COLOR24 | ILC_MASK, 16, 0);
    // kMask: light grey used as the transparent/background colour for the mask.
    static constexpr COLORREF kMask = RGB(0xC0, 0xC0, 0xC0);
    auto makeTextIcon = [&](const wchar_t* label, COLORREF clr) {
        HDC hdcScreen = GetDC(nullptr);
        HDC memDC     = CreateCompatibleDC(hdcScreen);

        // Use a 24-bpp DIB section so the pixel data is well-defined (no
        // uninitialised alpha bytes that confuse ILC_COLOR32 image lists).
        BITMAPINFOHEADER bih{};
        bih.biSize        = sizeof(bih);
        bih.biWidth       = 20;
        bih.biHeight      = -20;  // top-down
        bih.biPlanes      = 1;
        bih.biBitCount    = 24;
        bih.biCompression = BI_RGB;
        void*   pBits = nullptr;
        HBITMAP hBmp  = CreateDIBSection(hdcScreen,
                            reinterpret_cast<BITMAPINFO*>(&bih),
                            DIB_RGB_COLORS, &pBits, nullptr, 0);
        if (hBmp) {
            HBITMAP hOld = reinterpret_cast<HBITMAP>(SelectObject(memDC, hBmp));

            RECT rc{ 0, 0, 20, 20 };
            HBRUSH hBg = CreateSolidBrush(kMask);
            FillRect(memDC, &rc, hBg);
            DeleteObject(hBg);

            SetTextColor(memDC, clr);
            SetBkMode(memDC, TRANSPARENT);

            HFONT hf = CreateFontW(14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                                   DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                   CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH, L"Segoe UI");
            HFONT hOldFont = reinterpret_cast<HFONT>(SelectObject(memDC, hf));
            DrawTextW(memDC, label, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            SelectObject(memDC, hOldFont);
            DeleteObject(hf);

            SelectObject(memDC, hOld);  // deselect before passing to image list
            ImageList_AddMasked(hFmtImg, hBmp, kMask);
            DeleteObject(hBmp);
        }
        DeleteDC(memDC);
        ReleaseDC(nullptr, hdcScreen);
    };
    // Build icon set: B I U S Al Ac Ar Aj In Out • # Fc Hi
    makeTextIcon(L"B",  RGB(0,0,0));       // 0 Bold
    makeTextIcon(L"I",  RGB(0,0,0));       // 1 Italic
    makeTextIcon(L"U",  RGB(0,0,0));       // 2 Underline
    makeTextIcon(L"S",  RGB(0,0,0));       // 3 Strikeout
    makeTextIcon(L"≡",  RGB(0,0,200));     // 4 Align left
    makeTextIcon(L"≡",  RGB(0,0,200));     // 5 Align center
    makeTextIcon(L"≡",  RGB(0,0,200));     // 6 Align right
    makeTextIcon(L"≡",  RGB(0,0,200));     // 7 Justify
    makeTextIcon(L"→",  RGB(0,120,0));     // 8 Indent
    makeTextIcon(L"←",  RGB(0,120,0));     // 9 Outdent
    makeTextIcon(L"•",  RGB(0,0,0));       // 10 Bullet
    makeTextIcon(L"1.", RGB(0,0,0));       // 11 Numbering
    makeTextIcon(L"A",  RGB(200,0,0));     // 12 Font color
    makeTextIcon(L"H",  RGB(200,180,0));   // 13 Highlight

    SendMessageW(m_hwndFmtTb, TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(hFmtImg));

    // Placeholder space for combos (276px)
    TBBUTTON fmtBtns[] = {
        {276, 0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0}, // space for combos
        {0,  0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {0,  ID_FORMAT_BOLD,        (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {1,  ID_FORMAT_ITALIC,      (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {2,  ID_FORMAT_UNDERLINE,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {3,  ID_FORMAT_STRIKEOUT,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {0,  0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {4,  ID_PARA_ALIGN_LEFT,    (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECKGROUP,  {}, 0, 0},
        {5,  ID_PARA_ALIGN_CENTER,  (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECKGROUP,  {}, 0, 0},
        {6,  ID_PARA_ALIGN_RIGHT,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECKGROUP,  {}, 0, 0},
        {7,  ID_PARA_ALIGN_JUSTIFY, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECKGROUP,  {}, 0, 0},
        {0,  0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {8,  ID_PARA_INDENT,        (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON,      {}, 0, 0},
        {9,  ID_PARA_OUTDENT,       (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON,      {}, 0, 0},
        {0,  0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {10, ID_PARA_BULLET,        (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {11, ID_PARA_NUMBERING,     (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_CHECK,       {}, 0, 0},
        {0,  0, (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_SEP, {}, 0, 0},
        {12, ID_FORMAT_FONTCOLOR,   (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON,      {}, 0, 0},
        {13, ID_FORMAT_BGCOLOR,     (BYTE)TBSTATE_ENABLED, (BYTE)BTNS_BUTTON,      {}, 0, 0},
    };
    SendMessageW(m_hwndFmtTb, TB_ADDBUTTONSW,
                 ARRAYSIZE(fmtBtns), reinterpret_cast<LPARAM>(fmtBtns));
    SendMessageW(m_hwndFmtTb, TB_AUTOSIZE, 0, 0);

    // Add both toolbars to rebar
    REBARBANDINFOW rbi{};
    rbi.cbSize  = sizeof(rbi);

    rbi.fMask   = RBBIM_CHILD | RBBIM_CHILDSIZE | RBBIM_STYLE | RBBIM_SIZE;
    rbi.fStyle  = RBBS_CHILDEDGE | RBBS_NOGRIPPER;
    rbi.hwndChild = m_hwndStdTb;
    SIZE szStd{}; SendMessageW(m_hwndStdTb, TB_GETMAXSIZE, 0,
                               reinterpret_cast<LPARAM>(&szStd));
    rbi.cyChild = szStd.cy; rbi.cyMinChild = szStd.cy;
    rbi.cx      = szStd.cx;
    SendMessageW(m_hwndRebar, RB_INSERTBANDW, static_cast<WPARAM>(-1),
                 reinterpret_cast<LPARAM>(&rbi));

    SIZE szFmt{}; SendMessageW(m_hwndFmtTb, TB_GETMAXSIZE, 0,
                               reinterpret_cast<LPARAM>(&szFmt));
    rbi.fStyle     = RBBS_CHILDEDGE | RBBS_NOGRIPPER | RBBS_BREAK; // own row
    rbi.hwndChild  = m_hwndFmtTb;
    rbi.cyChild    = szFmt.cy; rbi.cyMinChild = szFmt.cy;
    rbi.cx         = szFmt.cx;
    SendMessageW(m_hwndRebar, RB_INSERTBANDW, static_cast<WPARAM>(-1),
                 reinterpret_cast<LPARAM>(&rbi));
}

// ─────────────────────────────────────────────
// CreateStatusBar
// ─────────────────────────────────────────────
void MainWindow::CreateStatusBar() {
    m_hwndStatus = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0, m_hwnd,
        reinterpret_cast<HMENU>(IDC_STATUSBAR),
        Application::Instance().GetHInstance(), nullptr);

    int parts[] = { 180, 340, 520, -1 };
    SendMessageW(m_hwndStatus, SB_SETPARTS, 4, reinterpret_cast<LPARAM>(parts));
}

// ─────────────────────────────────────────────
// Font / Size combos
// ─────────────────────────────────────────────
void MainWindow::PopulateFontCombo() {
    LOGFONTW lf{};
    lf.lfCharSet = DEFAULT_CHARSET;
    HDC hdc = GetDC(m_hwnd);
    EnumFontFamiliesExW(hdc, &lf, [](const LOGFONTW* plf, const TEXTMETRICW*, DWORD, LPARAM lp) {
        if (plf->lfFaceName[0] != L'@') // skip vertical fonts
            SendMessageW(reinterpret_cast<HWND>(lp), CB_ADDSTRING, 0,
                         reinterpret_cast<LPARAM>(plf->lfFaceName));
        return 1;
    }, reinterpret_cast<LPARAM>(m_hwndFontCombo), 0);
    ReleaseDC(m_hwnd, hdc);

    SendMessageW(m_hwndFontCombo, CB_SELECTSTRING, static_cast<WPARAM>(-1),
                 reinterpret_cast<LPARAM>(L"맑은 고딕"));
}

void MainWindow::PopulateSizeCombo() {
    const wchar_t* sizes[] = {
        L"8",L"9",L"10",L"11",L"12",L"14",L"16",
        L"18",L"20",L"22",L"24",L"28",L"32",L"36",L"48",L"72"
    };
    for (auto* s : sizes)
        SendMessageW(m_hwndSizeCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(s));
    SendMessageW(m_hwndSizeCombo, CB_SELECTSTRING, static_cast<WPARAM>(-1),
                 reinterpret_cast<LPARAM>(L"11"));
}

// ─────────────────────────────────────────────
// Layout
// ─────────────────────────────────────────────
void MainWindow::OnSize(int cx, int cy) {
    // RB_GETBARHEIGHT returns the actual height the rebar needs based on its
    // bands, which is correct even before the rebar window has been explicitly
    // sized (the rebar may be at size 0,0 on first call).
    int rebarH = m_hwndRebar
        ? static_cast<int>(SendMessageW(m_hwndRebar, RB_GETBARHEIGHT, 0, 0))
        : 0;

    RECT rcStatus{};
    if (m_hwndStatus) GetWindowRect(m_hwndStatus, &rcStatus);
    int statusH = rcStatus.bottom - rcStatus.top;

    if (m_hwndRebar)
        SetWindowPos(m_hwndRebar, nullptr, 0, 0, cx, rebarH, SWP_NOZORDER);
    if (m_hwndStatus)
        SendMessageW(m_hwndStatus, WM_SIZE, 0, MAKELPARAM(cx, cy));

    int editorY = rebarH;
    int editorH = cy - rebarH - statusH;
    if (m_editor) m_editor->Move(0, editorY, cx, std::max(0, editorH));

    // Reposition font/size combos inside format toolbar
    if (m_hwndFontCombo) SetWindowPos(m_hwndFontCombo, nullptr, 2, 2, 200, 200, SWP_NOZORDER);
    if (m_hwndSizeCombo) SetWindowPos(m_hwndSizeCombo, nullptr, 206, 2, 54, 200, SWP_NOZORDER);
}

// ─────────────────────────────────────────────
// OnCommand — route all IDs
// ─────────────────────────────────────────────
void MainWindow::OnCommand(UINT id) {
    switch (id) {
    case ID_FILE_NEW:     FileNew();    break;
    case ID_FILE_OPEN:    FileOpen();   break;
    case ID_FILE_SAVE:    FileSave();   break;
    case ID_FILE_SAVEAS:  FileSaveAs(); break;
    case ID_FILE_PRINT:   FilePrint();  break;
    case ID_FILE_CLOSE:   FileClose();  break;
    case ID_FILE_EXIT:    SendMessageW(m_hwnd, WM_CLOSE, 0, 0); break;

    case ID_EDIT_UNDO:      m_editor->Undo();      break;
    case ID_EDIT_REDO:      m_editor->Redo();      break;
    case ID_EDIT_CUT:       m_editor->Cut();       break;
    case ID_EDIT_COPY:      m_editor->Copy();      break;
    case ID_EDIT_PASTE:     m_editor->Paste();     break;
    case ID_EDIT_SELECTALL: m_editor->SelectAll(); break;
    case ID_EDIT_FIND:      EditFind();    break;
    case ID_EDIT_REPLACE:   EditReplace(); break;
    case ID_EDIT_DOCINFO:   EditDocInfo(); break;

    case ID_INSERT_TABLE:     InsertTable();    break;
    case ID_INSERT_HYPERLINK: InsertHyperlink();break;
    case ID_INSERT_DATETIME:  InsertDatetime(); break;
    case ID_INSERT_CHARMAP:   InsertCharMap();  break;
    case ID_INSERT_COMMENT:   InsertComment();  break;
    case ID_INSERT_PICTURE:   InsertPicture();  break;
    case ID_INSERT_BOOKMARK:  InsertBookmark(); break;

    case ID_FORMAT_CHAR:      FormatChar();    break;
    case ID_FORMAT_PARA:      FormatPara();    break;
    case ID_FORMAT_COLUMNS:   FormatColumns(); break;
    case ID_FORMAT_BORDER:    FormatBorder();  break;
    case ID_FORMAT_STYLE:     FormatStyle();   break;
    case ID_FORMAT_BOLD:      m_editor->ToggleBold();      break;
    case ID_FORMAT_ITALIC:    m_editor->ToggleItalic();    break;
    case ID_FORMAT_UNDERLINE: m_editor->ToggleUnderline(); break;
    case ID_FORMAT_STRIKEOUT: m_editor->ToggleStrikeout(); break;
    case ID_FORMAT_SUPERSCRIPT: m_editor->ToggleSuperscript(); break;
    case ID_FORMAT_SUBSCRIPT:   m_editor->ToggleSubscript();   break;
    case ID_FORMAT_FONTCOLOR: {
        CHOOSECOLORW cc{};
        static COLORREF custColors[16]{};
        cc.lStructSize  = sizeof(cc);
        cc.hwndOwner    = m_hwnd;
        cc.lpCustColors = custColors;
        cc.rgbResult    = RGB(0,0,0);
        cc.Flags        = CC_FULLOPEN | CC_RGBINIT;
        if (ChooseColorW(&cc)) m_editor->SetFontColor(cc.rgbResult);
        break;
    }
    case ID_FORMAT_BGCOLOR: {
        CHOOSECOLORW cc{};
        static COLORREF custColors2[16]{};
        cc.lStructSize  = sizeof(cc);
        cc.hwndOwner    = m_hwnd;
        cc.lpCustColors = custColors2;
        cc.rgbResult    = RGB(255,255,0);
        cc.Flags        = CC_FULLOPEN | CC_RGBINIT;
        if (ChooseColorW(&cc)) m_editor->SetBgColor(cc.rgbResult);
        break;
    }

    case ID_PARA_ALIGN_LEFT:    m_editor->SetAlignment(PFA_LEFT);    break;
    case ID_PARA_ALIGN_CENTER:  m_editor->SetAlignment(PFA_CENTER);  break;
    case ID_PARA_ALIGN_RIGHT:   m_editor->SetAlignment(PFA_RIGHT);   break;
    case ID_PARA_ALIGN_JUSTIFY: m_editor->SetAlignment(PFA_JUSTIFY); break;
    case ID_PARA_INDENT:        m_editor->Indent(+1);  break;
    case ID_PARA_OUTDENT:       m_editor->Indent(-1);  break;
    case ID_PARA_BULLET:        m_editor->ToggleBullet();    break;
    case ID_PARA_NUMBERING:     m_editor->ToggleNumbering(); break;

    case ID_PAGE_SETUP:   PageSetup();  break;
    case ID_PAGE_HEADER:  PageHeader(); break;
    case ID_PAGE_FOOTER:  PageFooter(); break;
    case ID_PAGE_NUMBER:  PageNumber(); break;
    case ID_PAGE_SPLIT:   PageSplit();  break;

    case ID_SETTINGS_GRID:
        m_gridVisible = !m_gridVisible;
        CheckMenuItem(GetMenu(m_hwnd), ID_SETTINGS_GRID,
                      m_gridVisible ? MF_CHECKED : MF_UNCHECKED);
        InvalidateRect(m_hwnd, nullptr, TRUE);
        break;
    case ID_SETTINGS_USER:
    case ID_SETTINGS_LANGUAGE:
    case ID_SETTINGS_THEME:
        ShowSettings(); break;

    case ID_HELP_MANUAL:
        ShellExecuteW(m_hwnd, L"open",
                      L"https://github.com/elrang3843/whatsup/wiki",
                      nullptr, nullptr, SW_SHOWNORMAL);
        break;
    case ID_HELP_LICENSE:
        ShellExecuteW(m_hwnd, L"open",
                      L"https://www.apache.org/licenses/LICENSE-2.0",
                      nullptr, nullptr, SW_SHOWNORMAL);
        break;
    case ID_HELP_ABOUT: ShowAbout(); break;

    case IDC_COMBO_FONT: OnFontComboChange(); break;
    case IDC_COMBO_SIZE: OnSizeComboChange(); break;
    }
    UpdateToolbarState();
}

void MainWindow::OnNotify(NMHDR* nm) {
    if (!nm) return;
    // Toolbar tooltip text
    if (nm->code == TTN_GETDISPINFOW) {
        auto* ttdi = reinterpret_cast<NMTTDISPINFOW*>(nm);
        struct { UINT id; StrID sid; } tips[] = {
            {ID_FILE_NEW,    StrID::TT_NEW},    {ID_FILE_OPEN,   StrID::TT_OPEN},
            {ID_FILE_SAVE,   StrID::TT_SAVE},   {ID_FILE_PRINT,  StrID::TT_PRINT},
            {ID_EDIT_UNDO,   StrID::TT_UNDO},   {ID_EDIT_REDO,   StrID::TT_REDO},
            {ID_EDIT_CUT,    StrID::TT_CUT},    {ID_EDIT_COPY,   StrID::TT_COPY},
            {ID_EDIT_PASTE,  StrID::TT_PASTE},
            {ID_FORMAT_BOLD, StrID::TT_BOLD},   {ID_FORMAT_ITALIC,  StrID::TT_ITALIC},
            {ID_FORMAT_UNDERLINE,StrID::TT_UNDERLINE},
            {ID_FORMAT_STRIKEOUT,StrID::TT_STRIKEOUT},
            {ID_PARA_ALIGN_LEFT,StrID::TT_ALIGN_LEFT},
            {ID_PARA_ALIGN_CENTER,StrID::TT_ALIGN_CENTER},
            {ID_PARA_ALIGN_RIGHT,StrID::TT_ALIGN_RIGHT},
            {ID_PARA_ALIGN_JUSTIFY,StrID::TT_ALIGN_JUSTIFY},
            {ID_PARA_INDENT, StrID::TT_INDENT}, {ID_PARA_OUTDENT,StrID::TT_OUTDENT},
            {ID_FORMAT_FONTCOLOR,StrID::TT_FONT_COLOR},
            {ID_FORMAT_BGCOLOR,  StrID::TT_BGCOLOR},
            {ID_EDIT_FIND,   StrID::TT_FIND},   {ID_EDIT_REPLACE,StrID::TT_REPLACE},
        };
        for (auto& t : tips) {
            if (ttdi->hdr.idFrom == t.id) {
                ttdi->lpszText = const_cast<LPWSTR>(Localization::Get(t.sid));
                break;
            }
        }
    }
    // Font/Size combo notification (CBN_SELCHANGE forwarded via WM_COMMAND)
    if (nm->idFrom == IDC_COMBO_FONT && nm->code == CBN_SELCHANGE) OnFontComboChange();
    if (nm->idFrom == IDC_COMBO_SIZE && nm->code == CBN_SELCHANGE) OnSizeComboChange();
}

// ─────────────────────────────────────────────
// File operations
// ─────────────────────────────────────────────
void MainWindow::FileNew() {
    if (!ConfirmClose()) return;
    m_editor->Clear();
    ApplyTheme(); // restore theme colours after Clear's WM_SETTEXT
    m_doc->Reset();
    m_ac->ClearDictionary();
    UpdateTitleBar();
    UpdateStatusBar();
}

void MainWindow::FileOpen() {
    if (!ConfirmClose()) return;
    std::wstring filter = Localization::GetFileFilter();
    wchar_t path[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = m_hwnd;
    ofn.lpstrFilter = filter.c_str();
    ofn.lpstrFile   = path;
    ofn.nMaxFile    = MAX_PATH;
    ofn.Flags       = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    if (!GetOpenFileNameW(&ofn)) {
        // Non-zero extended error means dialog failed, not user-cancelled
        DWORD err = CommDlgExtendedError();
        if (err != 0) {
            wchar_t msg[64];
            swprintf_s(msg, L"파일 열기 대화상자 오류: 0x%04lX", err);
            MessageBoxW(m_hwnd, msg, L"오류", MB_ICONERROR);
        }
        return;
    }
    OpenFile(path);
}

bool MainWindow::OpenFile(const std::wstring& path) {
    auto result = FormatManager::Instance().Load(path, *m_doc);
    if (!result.ok) {
        MessageBoxW(m_hwnd, result.error.c_str(),
                    Localization::Get(StrID::APP_TITLE), MB_ICONERROR);
        return false;
    }
    if (result.rtf) m_editor->SetRtf(result.content.empty() ? "" :
        std::string(result.content.begin(), result.content.end()));
    else if (!result.content.empty())
        m_editor->SetText(result.content);

    // Reapply theme so text/background colours stay correct regardless of
    // what the RTF or SetText may have reset them to.
    ApplyTheme();

    m_doc->SetPath(path);
    m_doc->SetModified(false);
    m_editor->SetModified(false);
    m_ac->LearnText(m_editor->GetText());
    UpdateTitleBar();
    UpdateStatusBar();
    return true;
}

void MainWindow::FileSave() {
    if (!m_doc->HasPath()) { FileSaveAs(); return; }
    SaveFile(m_doc->Path());
}

void MainWindow::FileSaveAs() {
    std::wstring filter = Localization::GetFileFilter();
    wchar_t path[MAX_PATH] = {};
    if (m_doc->HasPath()) wcscpy_s(path, m_doc->Path().c_str());

    OPENFILENAMEW sfn{};
    sfn.lStructSize = sizeof(sfn);
    sfn.hwndOwner   = m_hwnd;
    sfn.lpstrFilter = filter.c_str();
    sfn.lpstrFile   = path;
    sfn.nMaxFile    = MAX_PATH;
    sfn.Flags       = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    sfn.lpstrDefExt = L"txt";
    if (GetSaveFileNameW(&sfn)) SaveFile(path);
}

bool MainWindow::SaveFile(const std::wstring& path) {
    // Warn before saving as TXT (lossy — all formatting is discarded)
    size_t dot = path.rfind(L'.');
    if (dot != std::wstring::npos) {
        std::wstring ext = path.substr(dot + 1);
        for (auto& c : ext) c = towlower(c);
        if (ext == L"txt") {
            int ans = MessageBoxW(m_hwnd,
                Localization::Get(StrID::MSG_TXT_LOSSY_SAVE),
                Localization::Get(StrID::MSG_TXT_LOSSY_TITLE),
                MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2);
            if (ans != IDYES) return false;
        }
    }

    std::wstring text = m_editor->GetText();
    std::string  rtf  = m_editor->GetRtf();
    m_doc->SetPath(path);

    auto result = FormatManager::Instance().Save(path, text, rtf, *m_doc);
    if (!result.ok) {
        MessageBoxW(m_hwnd, result.error.c_str(),
                    Localization::Get(StrID::APP_TITLE), MB_ICONERROR);
        return false;
    }
    m_doc->SetModified(false);
    m_editor->SetModified(false);
    UpdateTitleBar();
    return true;
}

void MainWindow::FilePrint() {
    m_editor->Print(m_hwnd, m_doc->DisplayTitle());
}

void MainWindow::FileClose() {
    if (!ConfirmClose()) return;
    m_editor->Clear();
    m_doc->Reset();
    UpdateTitleBar();
    UpdateStatusBar();
}

bool MainWindow::ConfirmClose() {
    if (!m_editor->IsModified()) return true;
    int ret = MessageBoxW(m_hwnd,
                          Localization::Get(StrID::MSG_UNSAVED_CHANGES),
                          Localization::Get(StrID::MSG_UNSAVED_TITLE),
                          MB_YESNOCANCEL | MB_ICONQUESTION);
    if (ret == IDCANCEL) return false;
    if (ret == IDYES)    FileSave();
    return true;
}

// ─────────────────────────────────────────────
// Edit
// ─────────────────────────────────────────────
void MainWindow::EditFind() { ShowFindDialog(false); }
void MainWindow::EditReplace() { ShowFindDialog(true); }

void MainWindow::ShowFindDialog(bool replace) {
    if (m_hwndFind && IsWindow(m_hwndFind)) {
        SetForegroundWindow(m_hwndFind);
        return;
    }
    m_pFindReplace = new FINDREPLACEW{};
    m_pFindReplace->lStructSize = sizeof(FINDREPLACEW);
    m_pFindReplace->hwndOwner   = m_hwnd;
    m_pFindReplace->Flags       = FR_DOWN;
    static wchar_t findBuf[512]{}, replaceBuf[512]{};
    m_pFindReplace->lpstrFindWhat   = findBuf;
    m_pFindReplace->wFindWhatLen    = 512;
    m_pFindReplace->lpstrReplaceWith= replaceBuf;
    m_pFindReplace->wReplaceWithLen = 512;

    m_hwndFind = replace ? ReplaceTextW(m_pFindReplace)
                         : FindTextW(m_pFindReplace);
}

void MainWindow::OnFindMsg(FINDREPLACEW* fr) {
    if (!fr) return;
    if (fr->Flags & FR_DIALOGTERM) {
        m_hwndFind = nullptr;
        delete m_pFindReplace;
        m_pFindReplace = nullptr;
        return;
    }
    Editor::FindOptions opts;
    opts.findText   = fr->lpstrFindWhat;
    opts.replaceText= fr->lpstrReplaceWith ? fr->lpstrReplaceWith : L"";
    opts.matchCase  = (fr->Flags & FR_MATCHCASE) != 0;
    opts.wholeWord  = (fr->Flags & FR_WHOLEWORD) != 0;
    opts.forward    = (fr->Flags & FR_DOWN) != 0;

    if (fr->Flags & FR_FINDNEXT)   m_editor->FindNext(opts);
    if (fr->Flags & FR_REPLACE)    { m_editor->FindNext(opts); m_editor->ReplaceSelection(opts.replaceText); }
    if (fr->Flags & FR_REPLACEALL) m_editor->ReplaceAll(opts);
}

void MainWindow::EditDocInfo() {
    auto stats = m_editor->ComputeStats();
    m_doc->Stats().words = stats.words;
    m_doc->Stats().chars = stats.chars;
    DocInfoDialog::Show(m_hwnd, *m_doc);
}

// ─────────────────────────────────────────────
// Insert
// ─────────────────────────────────────────────
void MainWindow::InsertTable() {
    TableOptions opts;
    if (InsertTableDialog::Show(m_hwnd, opts))
        m_editor->InsertTable(opts.rows, opts.cols, opts.hasHeader, opts.widthPct);
}

void MainWindow::InsertHyperlink() {
    HyperlinkOptions opts;
    opts.displayText = m_editor->GetSelText();
    if (InsertHyperlinkDialog::Show(m_hwnd, opts))
        m_editor->InsertHyperlink(opts.displayText, opts.url);
}

void MainWindow::InsertDatetime() {
    DateTimeOptions opts;
    if (InsertDateTimeDialog::Show(m_hwnd, opts))
        m_editor->InsertText(opts.format);
}

void MainWindow::InsertCharMap() {
    // Launch system character map
    ShellExecuteW(m_hwnd, L"open", L"charmap.exe", nullptr, nullptr, SW_SHOWNORMAL);
}

void MainWindow::InsertComment() {
    m_editor->InsertText(L"[주석: ]");
}

void MainWindow::InsertPicture() {
    wchar_t path[MAX_PATH] = {};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner   = m_hwnd;
    ofn.lpstrFilter = L"이미지 (*.png;*.jpg;*.bmp;*.gif)\0*.png;*.jpg;*.bmp;*.gif\0모든 파일\0*.*\0\0";
    ofn.lpstrFile   = path;
    ofn.nMaxFile    = MAX_PATH;
    ofn.Flags       = OFN_FILEMUSTEXIST;
    if (!GetOpenFileNameW(&ofn)) return;

    // Insert OLE image into Rich Edit
    HBITMAP hBmp = reinterpret_cast<HBITMAP>(
        LoadImageW(nullptr, path, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE));
    if (!hBmp) { MessageBoxW(m_hwnd, L"이미지를 로드할 수 없습니다.", L"오류", MB_ICONERROR); return; }
    // For a full OLE embed, use IRichEditOle::InsertObject — stub with bitmap copy
    OpenClipboard(m_hwnd);
    EmptyClipboard();
    SetClipboardData(CF_BITMAP, hBmp);
    CloseClipboard();
    m_editor->Paste();
    DeleteObject(hBmp);
}

void MainWindow::InsertBookmark() {
    // Show bookmark dialog (uses IDD_INSERT_BOOKMARK)
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_INSERT_BOOKMARK), m_hwnd, nullptr);
}

// ─────────────────────────────────────────────
// Format
// ─────────────────────────────────────────────
void MainWindow::FormatChar() {
    LOGFONTW lf{};
    CHOOSEFONTW cf{};
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner   = m_hwnd;
    cf.lpLogFont   = &lf;
    cf.Flags       = CF_SCREENFONTS | CF_EFFECTS | CF_INITTOLOGFONTSTRUCT;
    // Pre-fill from current selection
    CharFormat cur = m_editor->GetCharFormat(true);
    wcscpy_s(lf.lfFaceName, cur.fontName.empty() ? L"맑은 고딕" : cur.fontName.c_str());
    lf.lfHeight = -MulDiv(cur.fontSize / 2, GetDeviceCaps(GetDC(nullptr), LOGPIXELSY), 72);
    if (ChooseFontW(&cf)) {
        CharFormat nc{};
        nc.fontName  = lf.lfFaceName;
        nc.fontSize  = cf.iPointSize / 5; // decipoints → half-points
        nc.bold      = (lf.lfWeight >= FW_BOLD) ? 1 : 0;
        nc.italic    = lf.lfItalic   ? 1 : 0;
        nc.underline = lf.lfUnderline? 1 : 0;
        nc.strikeout = lf.lfStrikeOut? 1 : 0;
        nc.textColor = cf.rgbColors;
        m_editor->ApplyCharFormat(nc, true);
        // Update combos
        SendMessageW(m_hwndFontCombo, CB_SELECTSTRING, static_cast<WPARAM>(-1),
                     reinterpret_cast<LPARAM>(lf.lfFaceName));
        wchar_t szPt[8]; _snwprintf_s(szPt, _countof(szPt), _TRUNCATE, L"%d", cf.iPointSize / 10);
        SendMessageW(m_hwndSizeCombo, CB_SELECTSTRING, static_cast<WPARAM>(-1),
                     reinterpret_cast<LPARAM>(szPt));
    }
}

void MainWindow::FormatPara() {
    // Simple inline paragraph alignment via a quick message box menu
    // (full dialog would use IDD_PARAGRAPH — kept minimal here)
    HMENU hPop = CreatePopupMenu();
    AppendMenuW(hPop, MF_STRING, 1, L"왼쪽 정렬");
    AppendMenuW(hPop, MF_STRING, 2, L"가운데 정렬");
    AppendMenuW(hPop, MF_STRING, 3, L"오른쪽 정렬");
    AppendMenuW(hPop, MF_STRING, 4, L"양쪽 정렬");
    POINT pt{}; GetCursorPos(&pt);
    int choice = TrackPopupMenu(hPop, TPM_RETURNCMD | TPM_RIGHTBUTTON,
                                pt.x, pt.y, 0, m_hwnd, nullptr);
    DestroyMenu(hPop);
    WORD alignMap[] = { 0, PFA_LEFT, PFA_CENTER, PFA_RIGHT, PFA_JUSTIFY };
    if (choice >= 1 && choice <= 4) m_editor->SetAlignment(alignMap[choice]);
}

void MainWindow::FormatColumns() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_COLUMNS), m_hwnd, nullptr);
}
void MainWindow::FormatBorder() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_BORDER_SHADING), m_hwnd, nullptr);
}
void MainWindow::FormatStyle() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_STYLE), m_hwnd, nullptr);
}

// ─────────────────────────────────────────────
// Page
// ─────────────────────────────────────────────
void MainWindow::PageSetup() {
    PAGESETUPDLGW ps{};
    ps.lStructSize = sizeof(ps);
    ps.hwndOwner   = m_hwnd;
    ps.Flags       = PSD_MARGINS | PSD_INHUNDREDTHSOFMILLIMETERS;
    PageSetupDlgW(&ps);
}
void MainWindow::PageHeader() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_HEADER_FOOTER), m_hwnd, nullptr);
}
void MainWindow::PageFooter() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_HEADER_FOOTER), m_hwnd, nullptr);
}
void MainWindow::PageNumber() {
    DialogBoxW(Application::Instance().GetHInstance(),
               MAKEINTRESOURCEW(IDD_PAGE_NUMBER), m_hwnd, nullptr);
}
void MainWindow::PageSplit() { m_editor->InsertText(L"\f"); }

// ─────────────────────────────────────────────
// Settings / Help
// ─────────────────────────────────────────────
void MainWindow::ShowSettings() {
    auto& s = Application::Instance().Settings();
    if (SettingsDialog::Show(m_hwnd, s)) {
        Localization::SetLanguage(s.language);
        m_spell->SetLanguage(s.spellLanguage);
        Application::Instance().SaveSettings();
        BuildMenu(); // rebuild menu in new language
        ApplyTheme();
        UpdateStatusBar();
    }
}
void MainWindow::ShowHelp() {
    ShellExecuteW(m_hwnd, L"open",
                  L"https://github.com/elrang3843/whatsup/wiki",
                  nullptr, nullptr, SW_SHOWNORMAL);
}
void MainWindow::ShowLicense() {
    ShellExecuteW(m_hwnd, L"open",
                  L"https://www.apache.org/licenses/LICENSE-2.0",
                  nullptr, nullptr, SW_SHOWNORMAL);
}
void MainWindow::ShowAbout() { AboutDialog::Show(m_hwnd); }

// ─────────────────────────────────────────────
// Theme
// ─────────────────────────────────────────────
void MainWindow::ApplyTheme() {
    if (m_editor && m_editor->GetHwnd()) {
        m_editor->SetBackground(Application::Instance().BgColor());
        CharFormat cf{};
        cf.textColor = Application::Instance().TextColor();
        m_editor->ApplyCharFormat(cf, false);
    }
    InvalidateRect(m_hwnd, nullptr, TRUE);
}

// ─────────────────────────────────────────────
// Status / Title bar
// ─────────────────────────────────────────────
void MainWindow::UpdateStatusBar() {
    if (!m_hwndStatus) return;
    if (!m_editor || !m_editor->GetHwnd()) return;
    auto stats = m_editor->ComputeStats();
    m_doc->Stats().words = stats.words;
    m_doc->Stats().chars = stats.chars;

    wchar_t buf[128];
    // Part 0: page (approximate)
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"페이지: 1 / 1");
    SendMessageW(m_hwndStatus, SB_SETTEXTW, 0, reinterpret_cast<LPARAM>(buf));
    // Part 1: word count
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"단어 수: %d", stats.words);
    SendMessageW(m_hwndStatus, SB_SETTEXTW, 1, reinterpret_cast<LPARAM>(buf));
    // Part 2: language
    const wchar_t* langStr =
        (Application::Instance().Settings().language == Language::Korean)
        ? L"한국어 / English" : L"English / 한국어";
    SendMessageW(m_hwndStatus, SB_SETTEXTW, 2, reinterpret_cast<LPARAM>(langStr));
    // Part 3: line / col
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"줄: %d | 칸: %d",
                 m_editor->GetCaretLine(), m_editor->GetCaretCol());
    SendMessageW(m_hwndStatus, SB_SETTEXTW, 3, reinterpret_cast<LPARAM>(buf));
}

void MainWindow::UpdateTitleBar() {
    std::wstring title = m_doc->DisplayTitle() + L" - " +
                         Localization::Get(StrID::APP_TITLE);
    SetWindowTextW(m_hwnd, title.c_str());
}

void MainWindow::UpdateToolbarState() {
    if (!m_editor || !m_editor->GetHwnd()) return;
    auto en = [&](HWND tb, UINT id, bool e) {
        SendMessageW(tb, TB_ENABLEBUTTON, id, e ? TRUE : FALSE);
    };
    en(m_hwndStdTb, ID_EDIT_UNDO,  m_editor->CanUndo());
    en(m_hwndStdTb, ID_EDIT_REDO,  m_editor->CanRedo());
    en(m_hwndStdTb, ID_EDIT_PASTE, m_editor->CanPaste());

    // Sync bold/italic/underline check state
    auto chk = [&](UINT id, bool c) {
        SendMessageW(m_hwndFmtTb, TB_CHECKBUTTON, id, c ? TRUE : FALSE);
    };
    CharFormat cf = m_editor->GetCharFormat(true);
    chk(ID_FORMAT_BOLD,      cf.bold      == 1);
    chk(ID_FORMAT_ITALIC,    cf.italic    == 1);
    chk(ID_FORMAT_UNDERLINE, cf.underline == 1);
    chk(ID_FORMAT_STRIKEOUT, cf.strikeout == 1);

    ParaFormat pf = m_editor->GetParaFormat();
    chk(ID_PARA_ALIGN_LEFT,    pf.alignment == PFA_LEFT);
    chk(ID_PARA_ALIGN_CENTER,  pf.alignment == PFA_CENTER);
    chk(ID_PARA_ALIGN_RIGHT,   pf.alignment == PFA_RIGHT);
    chk(ID_PARA_ALIGN_JUSTIFY, pf.alignment == PFA_JUSTIFY);
}

// ─────────────────────────────────────────────
// Text change → spell / autocomplete / stats
// ─────────────────────────────────────────────
void MainWindow::OnTextChanged() {
    m_doc->SetModified(true);
    m_editor->SetModified(true);
    UpdateTitleBar();
    UpdateStatusBar();

    auto& s = Application::Instance().Settings();
    if (s.autocomplete) {
        std::wstring text = m_editor->GetText();
        m_ac->OnTextChanged(text, m_editor->GetCaretPos());
    }

    // Lightweight spell check on current word only
    if (s.spellEnabled && m_spell->IsAvailable()) {
        // (Full document pass is deferred to avoid per-keystroke lag)
    }
}

// ─────────────────────────────────────────────
// Font / Size combo changes
// ─────────────────────────────────────────────
void MainWindow::OnFontComboChange() {
    wchar_t name[LF_FACESIZE] = {};
    int idx = static_cast<int>(SendMessageW(m_hwndFontCombo, CB_GETCURSEL, 0, 0));
    if (idx < 0) return;
    SendMessageW(m_hwndFontCombo, CB_GETLBTEXT, idx, reinterpret_cast<LPARAM>(name));
    CharFormat cf{}; cf.fontName = name;
    m_editor->ApplyCharFormat(cf, true);
    SetFocus(m_editor->GetHwnd());
}

void MainWindow::OnSizeComboChange() {
    wchar_t buf[16] = {};
    GetWindowTextW(m_hwndSizeCombo, buf, _countof(buf));
    int pt = _wtoi(buf);
    if (pt <= 0) return;
    CharFormat cf{}; cf.fontSize = pt * 2; // points → half-points
    m_editor->ApplyCharFormat(cf, true);
    SetFocus(m_editor->GetHwnd());
}
