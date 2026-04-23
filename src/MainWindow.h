#pragma once
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <string>
#include <memory>
#include "Editor.h"
#include "Document.h"
#include "spell/SpellChecker.h"
#include "autocomplete/AutoComplete.h"

class MainWindow {
public:
    static HWND Create(HINSTANCE hInst);
    static HWND GetFindDlg() { return s_instance ? s_instance->m_hwndFind : nullptr; }

private:
    // Window proc
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);

    // Creation
    bool OnCreate(HWND hwnd, HINSTANCE hInst);
    void BuildMenu();
    void CreateToolbars();
    void CreateStatusBar();
    void PopulateFontCombo();
    void PopulateSizeCombo();

    // Layout
    void OnSize(int cx, int cy);

    // Commands
    void OnCommand(UINT id);
    void OnNotify(NMHDR* nmhdr);

    // File operations
    void FileNew();
    void FileOpen();
    void FileSave();
    void FileSaveAs();
    void FilePrint();
    void FileClose();

    // Edit operations
    void EditFind();
    void EditReplace();
    void EditDocInfo();

    // Insert
    void InsertTable();
    void InsertHyperlink();
    void InsertDatetime();
    void InsertCharMap();
    void InsertComment();
    void InsertPicture();
    void InsertBookmark();

    // Format
    void FormatChar();
    void FormatPara();
    void FormatColumns();
    void FormatBorder();
    void FormatStyle();

    // Page
    void PageSetup();
    void PageHeader();
    void PageFooter();
    void PageNumber();
    void PageSplit();

    // Settings
    void ShowSettings();

    // Help
    void ShowHelp();
    void ShowLicense();
    void ShowAbout();

    // Theme application
    void ApplyTheme();

    // Status bar helpers
    void UpdateStatusBar();
    void UpdateTitleBar();
    void UpdateToolbarState();

    // Spell check / text changes
    void OnTextChanged();

    // Open file given path (drag-drop / command line)
    bool OpenFile(const std::wstring& path);
    bool SaveFile(const std::wstring& path);

    // Check for unsaved changes; returns false if user cancels
    bool ConfirmClose();

    // Font/Size combo changed
    void OnFontComboChange();
    void OnSizeComboChange();

    // Find/Replace common dialog
    void ShowFindDialog(bool replace);
    static UINT s_findMsgId;          // registered FindReplace message
    void OnFindMsg(FINDREPLACEW* fr);

    // ---- Member variables ----
    HWND            m_hwnd          = nullptr;
    HWND            m_hwndRebar     = nullptr;
    HWND            m_hwndStdTb     = nullptr;
    HWND            m_hwndFmtTb     = nullptr;
    HWND            m_hwndStatus    = nullptr;
    HWND            m_hwndFontCombo = nullptr;
    HWND            m_hwndSizeCombo = nullptr;

    std::unique_ptr<Editor>        m_editor;
    std::unique_ptr<Document>      m_doc;
    std::unique_ptr<SpellChecker>  m_spell;
    std::unique_ptr<AutoComplete>  m_ac;

    FINDREPLACEW*   m_pFindReplace  = nullptr;
    HWND            m_hwndFind      = nullptr; // modeless find/replace dialog

    bool            m_gridVisible   = false;
    bool            m_fullyCreated  = false;

    // Toolbar image list (standard icons)
    HIMAGELIST      m_hImgList      = nullptr;

    static MainWindow* s_instance;

    // Window class name
    static constexpr wchar_t kClassName[] = L"WhatsUpMainWnd";
};
