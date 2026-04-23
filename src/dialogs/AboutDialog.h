#pragma once
#include <windows.h>

class AboutDialog {
public:
    static void Show(HWND hwndParent);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
};
