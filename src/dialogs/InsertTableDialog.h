#pragma once
#include <windows.h>

struct TableOptions {
    int  rows       = 3;
    int  cols       = 3;
    bool autoFit    = true;
    bool hasHeader  = true;
    int  widthPct   = 100;
};

class InsertTableDialog {
public:
    static bool Show(HWND hwndParent, TableOptions& opts);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
};
