#pragma once
#include <windows.h>
#include <string>

struct DateTimeOptions {
    std::wstring format;
    bool         autoUpdate = false;
};

class InsertDateTimeDialog {
public:
    static bool Show(HWND hwndParent, DateTimeOptions& opts);

private:
    static INT_PTR CALLBACK DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void PopulateFormats(HWND hwndList);
    static std::wstring FormatDateTime(const std::wstring& fmt);
};
