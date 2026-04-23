#pragma once
#include <windows.h>
#include "../Application.h"

class SettingsDialog {
public:
    // Returns true if settings were changed
    static bool Show(HWND hwndParent, AppSettings& settings);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void Populate(HWND hwnd, const AppSettings& settings);
    static void Collect (HWND hwnd,       AppSettings& settings);
};
