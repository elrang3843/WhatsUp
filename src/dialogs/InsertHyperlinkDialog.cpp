#include "InsertHyperlinkDialog.h"
#include "../resource.h"

bool InsertHyperlinkDialog::Show(HWND hwndParent, HyperlinkOptions& opts) {
    return DialogBoxParamW(GetModuleHandleW(nullptr),
                           MAKEINTRESOURCEW(IDD_INSERT_HYPERLINK),
                           hwndParent, DlgProc,
                           reinterpret_cast<LPARAM>(&opts)) == IDOK;
}

INT_PTR CALLBACK InsertHyperlinkDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static HyperlinkOptions* pOpts = nullptr;
    wchar_t buf[1024];

    switch (msg) {
    case WM_INITDIALOG:
        pOpts = reinterpret_cast<HyperlinkOptions*>(lParam);
        SetDlgItemTextW(hwnd, IDC_LINK_DISPLAY, pOpts->displayText.c_str());
        SetDlgItemTextW(hwnd, IDC_LINK_URL,     pOpts->url.c_str());
        SetDlgItemTextW(hwnd, IDC_LINK_TOOLTIP, pOpts->tooltip.c_str());
        return TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK && pOpts) {
            GetDlgItemTextW(hwnd, IDC_LINK_DISPLAY, buf, _countof(buf));
            pOpts->displayText = buf;
            GetDlgItemTextW(hwnd, IDC_LINK_URL,     buf, _countof(buf));
            pOpts->url = buf;
            GetDlgItemTextW(hwnd, IDC_LINK_TOOLTIP, buf, _countof(buf));
            pOpts->tooltip = buf;
            EndDialog(hwnd, IDOK);
            return TRUE;
        }
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hwnd, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}
