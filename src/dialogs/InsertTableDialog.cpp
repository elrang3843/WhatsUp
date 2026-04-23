#include "InsertTableDialog.h"
#include "../resource.h"

bool InsertTableDialog::Show(HWND hwndParent, TableOptions& opts) {
    return DialogBoxParamW(GetModuleHandleW(nullptr),
                           MAKEINTRESOURCEW(IDD_INSERT_TABLE),
                           hwndParent, DlgProc,
                           reinterpret_cast<LPARAM>(&opts)) == IDOK;
}

INT_PTR CALLBACK InsertTableDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static TableOptions* pOpts = nullptr;
    switch (msg) {
    case WM_INITDIALOG:
        pOpts = reinterpret_cast<TableOptions*>(lParam);
        SetDlgItemInt(hwnd, IDC_TABLE_ROWS,      pOpts->rows,      FALSE);
        SetDlgItemInt(hwnd, IDC_TABLE_COLS,      pOpts->cols,      FALSE);
        SetDlgItemInt(hwnd, IDC_TABLE_WIDTH_PCT, pOpts->widthPct,  FALSE);
        CheckDlgButton(hwnd, IDC_TABLE_AUTOFIT,   pOpts->autoFit   ? BST_CHECKED : BST_UNCHECKED);
        CheckDlgButton(hwnd, IDC_TABLE_HAS_HEADER,pOpts->hasHeader ? BST_CHECKED : BST_UNCHECKED);
        return TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            if (pOpts) {
                BOOL ok1, ok2;
                pOpts->rows      = GetDlgItemInt(hwnd, IDC_TABLE_ROWS,       &ok1, FALSE);
                pOpts->cols      = GetDlgItemInt(hwnd, IDC_TABLE_COLS,       &ok2, FALSE);
                pOpts->widthPct  = GetDlgItemInt(hwnd, IDC_TABLE_WIDTH_PCT,  &ok1, FALSE);
                pOpts->autoFit   = IsDlgButtonChecked(hwnd, IDC_TABLE_AUTOFIT)    == BST_CHECKED;
                pOpts->hasHeader = IsDlgButtonChecked(hwnd, IDC_TABLE_HAS_HEADER) == BST_CHECKED;
                if (pOpts->rows < 1) pOpts->rows = 1;
                if (pOpts->cols < 1) pOpts->cols = 1;
            }
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
