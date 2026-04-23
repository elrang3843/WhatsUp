#include "DocInfoDialog.h"
#include "../resource.h"
#include "../i18n/Localization.h"
#include <cstdio>

bool DocInfoDialog::Show(HWND hwndParent, Document& doc) {
    return DialogBoxParamW(GetModuleHandleW(nullptr),
                           MAKEINTRESOURCEW(IDD_DOC_INFO),
                           hwndParent, DlgProc,
                           reinterpret_cast<LPARAM>(&doc)) == IDOK;
}

void DocInfoDialog::Populate(HWND hwnd, const Document& doc) {
    const auto& p = doc.Properties();
    SetDlgItemTextW(hwnd, IDC_DOC_TITLE,    p.title.c_str());
    SetDlgItemTextW(hwnd, IDC_DOC_AUTHOR,   p.author.c_str());
    SetDlgItemTextW(hwnd, IDC_DOC_SUBJECT,  p.subject.c_str());
    SetDlgItemTextW(hwnd, IDC_DOC_KEYWORDS, p.keywords.c_str());
    SetDlgItemTextW(hwnd, IDC_DOC_COMMENT,  p.comment.c_str());

    const auto& st = doc.Stats();
    wchar_t buf[64];
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"%d", st.words);
    SetDlgItemTextW(hwnd, IDC_WORD_COUNT, buf);
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"%d", st.chars);
    SetDlgItemTextW(hwnd, IDC_CHAR_COUNT, buf);
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"%d", st.pages);
    SetDlgItemTextW(hwnd, IDC_PAGE_COUNT, buf);

    // Watermark
    const auto& wm = p.watermark;
    CheckRadioButton(hwnd, IDC_WATERMARK_NONE, IDC_WATERMARK_IMAGE,
        wm.type == WatermarkType::None  ? IDC_WATERMARK_NONE  :
        wm.type == WatermarkType::Text  ? IDC_WATERMARK_TEXT  : IDC_WATERMARK_IMAGE);
    SetDlgItemTextW(hwnd, IDC_WATERMARK_CONTENT,
        wm.type == WatermarkType::Text  ? wm.text.c_str() :
        wm.type == WatermarkType::Image ? wm.imagePath.c_str() : L"");

    UpdateWatermarkControls(hwnd);
}

void DocInfoDialog::UpdateWatermarkControls(HWND hwnd) {
    bool hasWm = IsDlgButtonChecked(hwnd, IDC_WATERMARK_NONE) != BST_CHECKED;
    EnableWindow(GetDlgItem(hwnd, IDC_WATERMARK_CONTENT), hasWm);
}

void DocInfoDialog::Collect(HWND hwnd, Document& doc) {
    auto& p = doc.Properties();
    wchar_t buf[512];
    GetDlgItemTextW(hwnd, IDC_DOC_TITLE,    buf, _countof(buf)); p.title    = buf;
    GetDlgItemTextW(hwnd, IDC_DOC_AUTHOR,   buf, _countof(buf)); p.author   = buf;
    GetDlgItemTextW(hwnd, IDC_DOC_SUBJECT,  buf, _countof(buf)); p.subject  = buf;
    GetDlgItemTextW(hwnd, IDC_DOC_KEYWORDS, buf, _countof(buf)); p.keywords = buf;
    GetDlgItemTextW(hwnd, IDC_DOC_COMMENT,  buf, _countof(buf)); p.comment  = buf;

    auto& wm = p.watermark;
    if (IsDlgButtonChecked(hwnd, IDC_WATERMARK_TEXT) == BST_CHECKED) {
        wm.type = WatermarkType::Text;
        GetDlgItemTextW(hwnd, IDC_WATERMARK_CONTENT, buf, _countof(buf));
        wm.text = buf;
    } else if (IsDlgButtonChecked(hwnd, IDC_WATERMARK_IMAGE) == BST_CHECKED) {
        wm.type = WatermarkType::Image;
        GetDlgItemTextW(hwnd, IDC_WATERMARK_CONTENT, buf, _countof(buf));
        wm.imagePath = buf;
    } else {
        wm.type = WatermarkType::None;
    }
}

INT_PTR CALLBACK DocInfoDialog::DlgProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static Document* pDoc = nullptr;
    switch (msg) {
    case WM_INITDIALOG:
        pDoc = reinterpret_cast<Document*>(lParam);
        Populate(hwnd, *pDoc);
        return TRUE;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_WATERMARK_NONE:
        case IDC_WATERMARK_TEXT:
        case IDC_WATERMARK_IMAGE:
            UpdateWatermarkControls(hwnd);
            return TRUE;
        case IDOK:
            if (pDoc) Collect(hwnd, *pDoc);
            EndDialog(hwnd, IDOK);
            return TRUE;
        case IDCANCEL:
            EndDialog(hwnd, IDCANCEL);
            return TRUE;
        }
        break;
    }
    return FALSE;
}
