#include "InsertDateTimeDialog.h"
#include "../resource.h"

// Korean/mixed date-time format strings
static const wchar_t* const s_formats[] = {
    L"yyyy-MM-dd",
    L"yyyy년 MM월 dd일",
    L"MM/dd/yyyy",
    L"dd/MM/yyyy",
    L"yyyy-MM-dd HH:mm:ss",
    L"yyyy년 MM월 dd일 HH시 mm분",
    L"HH:mm:ss",
    L"HH:mm",
    L"h:mm tt",          // 12-hour
    L"dddd, MMMM d, yyyy",
    L"MMMM d, yyyy",
    L"d MMMM yyyy",
};

std::wstring InsertDateTimeDialog::FormatDateTime(const std::wstring& fmt) {
    SYSTEMTIME st{};
    GetLocalTime(&st);

    // Simple format substitution
    std::wstring r = fmt;
    auto rep = [&](const wchar_t* from, auto valFn) {
        std::wstring f(from);
        wchar_t buf[16];
        size_t pos = 0;
        while ((pos = r.find(f, pos)) != std::wstring::npos) {
            valFn(buf, _countof(buf));
            r.replace(pos, f.size(), buf);
            pos += wcslen(buf);
        }
    };

    rep(L"yyyy", [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%04d", st.wYear); });
    rep(L"MM",   [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%02d", st.wMonth); });
    rep(L"dd",   [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%02d", st.wDay); });
    rep(L"HH",   [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%02d", st.wHour); });
    rep(L"mm",   [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%02d", st.wMinute); });
    rep(L"ss",   [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%02d", st.wSecond); });

    // Day/month names (Windows locale)
    static const wchar_t* days[]   = { L"일요일",L"월요일",L"화요일",L"수요일",L"목요일",L"금요일",L"토요일" };
    static const wchar_t* months[] = { L"",L"1월",L"2월",L"3월",L"4월",L"5월",L"6월",
                                       L"7월",L"8월",L"9월",L"10월",L"11월",L"12월" };
    {
        std::wstring f(L"dddd");
        size_t pos = 0;
        while ((pos = r.find(f, pos)) != std::wstring::npos) {
            r.replace(pos, f.size(), days[st.wDayOfWeek]);
            pos += wcslen(days[st.wDayOfWeek]);
        }
    }
    {
        std::wstring f(L"MMMM");
        size_t pos = 0;
        while ((pos = r.find(f, pos)) != std::wstring::npos) {
            r.replace(pos, f.size(), months[st.wMonth]);
            pos += wcslen(months[st.wMonth]);
        }
    }
    rep(L"d",  [&](wchar_t* b, int n) { _snwprintf_s(b, n, _TRUNCATE, L"%d", st.wDay); });
    rep(L"년", [](wchar_t* b, int) { wcscpy_s(b, 4, L"년"); });
    rep(L"월", [](wchar_t* b, int) { wcscpy_s(b, 4, L"월"); });
    rep(L"일", [](wchar_t* b, int) { wcscpy_s(b, 4, L"일"); });
    rep(L"시", [](wchar_t* b, int) { wcscpy_s(b, 4, L"시"); });
    rep(L"분", [](wchar_t* b, int) { wcscpy_s(b, 4, L"분"); });
    return r;
}

void InsertDateTimeDialog::PopulateFormats(HWND hwndList) {
    SendMessageW(hwndList, LB_RESETCONTENT, 0, 0);
    for (const auto* fmt : s_formats) {
        std::wstring preview = FormatDateTime(fmt);
        SendMessageW(hwndList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(preview.c_str()));
    }
}

bool InsertDateTimeDialog::Show(HWND hwndParent, DateTimeOptions& opts) {
    return DialogBoxParamW(GetModuleHandleW(nullptr),
                           MAKEINTRESOURCEW(IDD_INSERT_DATETIME),
                           hwndParent, DlgProc,
                           reinterpret_cast<LPARAM>(&opts)) == IDOK;
}

INT_PTR CALLBACK InsertDateTimeDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static DateTimeOptions* pOpts = nullptr;
    switch (msg) {
    case WM_INITDIALOG:
        pOpts = reinterpret_cast<DateTimeOptions*>(lParam);
        PopulateFormats(GetDlgItem(hwnd, IDC_DATETIME_LIST));
        SendMessageW(GetDlgItem(hwnd, IDC_DATETIME_LIST), LB_SETCURSEL, 0, 0);
        {
            std::wstring preview = FormatDateTime(s_formats[0]);
            SetDlgItemTextW(hwnd, IDC_DATETIME_PREVIEW, preview.c_str());
        }
        CheckDlgButton(hwnd, IDC_DATETIME_AUTO,
                       pOpts->autoUpdate ? BST_CHECKED : BST_UNCHECKED);
        return TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDC_DATETIME_LIST && HIWORD(wParam) == LBN_SELCHANGE) {
            int idx = static_cast<int>(
                SendMessageW(GetDlgItem(hwnd, IDC_DATETIME_LIST), LB_GETCURSEL, 0, 0));
            if (idx >= 0 && idx < static_cast<int>(ARRAYSIZE(s_formats))) {
                std::wstring preview = FormatDateTime(s_formats[idx]);
                SetDlgItemTextW(hwnd, IDC_DATETIME_PREVIEW, preview.c_str());
            }
            return TRUE;
        }
        if (LOWORD(wParam) == IDOK && pOpts) {
            int idx = static_cast<int>(
                SendMessageW(GetDlgItem(hwnd, IDC_DATETIME_LIST), LB_GETCURSEL, 0, 0));
            if (idx >= 0 && idx < static_cast<int>(ARRAYSIZE(s_formats)))
                pOpts->format = FormatDateTime(s_formats[idx]);
            pOpts->autoUpdate = IsDlgButtonChecked(hwnd, IDC_DATETIME_AUTO) == BST_CHECKED;
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
