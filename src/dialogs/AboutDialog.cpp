#include "AboutDialog.h"
#include "../resource.h"
#include "../i18n/Localization.h"
#include <shellapi.h>

void AboutDialog::Show(HWND hwndParent) {
    DialogBoxW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_ABOUT),
               hwndParent, DlgProc);
}

INT_PTR CALLBACK AboutDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_INITDIALOG: {
        // Set version and copyright text
        std::wstring ver = std::wstring(L"Version ") + Localization::Get(StrID::APP_VERSION);
        SetDlgItemTextW(hwnd, IDC_ABOUT_VERSION,   ver.c_str());
        SetDlgItemTextW(hwnd, IDC_ABOUT_COPYRIGHT, Localization::Get(StrID::APP_COPYRIGHT));

        // Make URL clickable (SysLink alternative: just set as static)
        SetDlgItemTextW(hwnd, IDC_ABOUT_URL,
                        L"https://github.com/elrang3843/whatsup");

        // Update caption in current language
        SetWindowTextW(hwnd, Localization::Get(StrID::MENU_HELP_ABOUT));
        return TRUE;
    }
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hwnd, LOWORD(wParam));
            return TRUE;
        }
        // Open URL if user clicks on IDC_ABOUT_URL label
        if (LOWORD(wParam) == IDC_ABOUT_URL) {
            ShellExecuteW(hwnd, L"open",
                          L"https://github.com/elrang3843/whatsup",
                          nullptr, nullptr, SW_SHOWNORMAL);
            return TRUE;
        }
        break;
    case WM_NOTIFY:
        // Handle SysLink clicks if IDC_ABOUT_URL is a SysLink control
        break;
    }
    return FALSE;
}
