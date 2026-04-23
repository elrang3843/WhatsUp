#include "SettingsDialog.h"
#include "../resource.h"
#include "../i18n/Localization.h"

bool SettingsDialog::Show(HWND hwndParent, AppSettings& settings) {
    return DialogBoxParamW(GetModuleHandleW(nullptr),
                           MAKEINTRESOURCEW(IDD_SETTINGS),
                           hwndParent, DlgProc,
                           reinterpret_cast<LPARAM>(&settings)) == IDOK;
}

void SettingsDialog::Populate(HWND hwnd, const AppSettings& s) {
    // Language combo
    HWND hLang = GetDlgItem(hwnd, IDC_LANG_COMBO);
    SendMessageW(hLang, CB_RESETCONTENT, 0, 0);
    SendMessageW(hLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"한국어 (Korean)"));
    SendMessageW(hLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"English"));
    SendMessageW(hLang, CB_SETCURSEL,
                 s.language == Language::Korean ? 0 : 1, 0);

    // Theme combo
    HWND hTheme = GetDlgItem(hwnd, IDC_THEME_COMBO);
    SendMessageW(hTheme, CB_RESETCONTENT, 0, 0);
    SendMessageW(hTheme, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(
        Localization::Get(SID::TAB_GENERAL))); // reuse for "Light"
    SendMessageW(hTheme, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"Dark"));
    SendMessageW(hTheme, CB_SETCURSEL, s.theme == Theme::Dark ? 1 : 0, 0);

    SetDlgItemTextW(hwnd, IDC_USER_NAME, s.userName.c_str());
    SetDlgItemTextW(hwnd, IDC_USER_ORG,  s.userOrg.c_str());

    CheckDlgButton(hwnd, IDC_SPELL_ENABLE,        s.spellEnabled  ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hwnd, IDC_AUTOCOMPLETE_ENABLE, s.autocomplete   ? BST_CHECKED : BST_UNCHECKED);

    HWND hSpellLang = GetDlgItem(hwnd, IDC_SPELL_LANG);
    SendMessageW(hSpellLang, CB_RESETCONTENT, 0, 0);
    SendMessageW(hSpellLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"ko-KR"));
    SendMessageW(hSpellLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"en-US"));
    SendMessageW(hSpellLang, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L"en-GB"));
    SendMessageW(hSpellLang, CB_SELECTSTRING, static_cast<WPARAM>(-1),
                 reinterpret_cast<LPARAM>(s.spellLanguage.c_str()));

    CheckDlgButton(hwnd, IDC_GRID_SHOW, s.gridShow ? BST_CHECKED : BST_UNCHECKED);
    CheckDlgButton(hwnd, IDC_GRID_SNAP, s.gridSnap ? BST_CHECKED : BST_UNCHECKED);
}

void SettingsDialog::Collect(HWND hwnd, AppSettings& s) {
    HWND hLang  = GetDlgItem(hwnd, IDC_LANG_COMBO);
    int  langIdx = static_cast<int>(SendMessageW(hLang, CB_GETCURSEL, 0, 0));
    s.language = (langIdx == 1) ? Language::English : Language::Korean;

    HWND hTheme  = GetDlgItem(hwnd, IDC_THEME_COMBO);
    int  themeIdx = static_cast<int>(SendMessageW(hTheme, CB_GETCURSEL, 0, 0));
    s.theme = (themeIdx == 1) ? Theme::Dark : Theme::Light;

    wchar_t buf[256];
    GetDlgItemTextW(hwnd, IDC_USER_NAME, buf, _countof(buf)); s.userName = buf;
    GetDlgItemTextW(hwnd, IDC_USER_ORG,  buf, _countof(buf)); s.userOrg  = buf;

    s.spellEnabled = IsDlgButtonChecked(hwnd, IDC_SPELL_ENABLE)        == BST_CHECKED;
    s.autocomplete = IsDlgButtonChecked(hwnd, IDC_AUTOCOMPLETE_ENABLE) == BST_CHECKED;

    HWND hSpellLang = GetDlgItem(hwnd, IDC_SPELL_LANG);
    int  slIdx = static_cast<int>(SendMessageW(hSpellLang, CB_GETCURSEL, 0, 0));
    if (slIdx >= 0) {
        wchar_t slBuf[64];
        SendMessageW(hSpellLang, CB_GETLBTEXT, slIdx, reinterpret_cast<LPARAM>(slBuf));
        s.spellLanguage = slBuf;
    }

    s.gridShow = IsDlgButtonChecked(hwnd, IDC_GRID_SHOW) == BST_CHECKED;
    s.gridSnap = IsDlgButtonChecked(hwnd, IDC_GRID_SNAP) == BST_CHECKED;
}

INT_PTR CALLBACK SettingsDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static AppSettings* pSettings = nullptr;

    switch (msg) {
    case WM_INITDIALOG:
        pSettings = reinterpret_cast<AppSettings*>(lParam);
        Populate(hwnd, *pSettings);
        return TRUE;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDOK:
            if (pSettings) Collect(hwnd, *pSettings);
            EndDialog(hwnd, IDOK);
            return TRUE;
        case IDCANCEL:
            EndDialog(hwnd, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}
