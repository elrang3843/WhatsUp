#pragma once
#include <windows.h>
#include <string>

struct HyperlinkOptions {
    std::wstring displayText;
    std::wstring url;
    std::wstring tooltip;
};

class InsertHyperlinkDialog {
public:
    static bool Show(HWND hwndParent, HyperlinkOptions& opts);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
};
