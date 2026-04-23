#pragma once
#include <windows.h>
#include "../Document.h"

class DocInfoDialog {
public:
    static bool Show(HWND hwndParent, Document& doc);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void Populate(HWND hwnd, const Document& doc);
    static void Collect (HWND hwnd,       Document& doc);
    static void UpdateWatermarkControls(HWND hwnd);
};
