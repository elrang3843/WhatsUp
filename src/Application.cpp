#include "Application.h"
#include "MainWindow.h"
#include "resource.h"
#include <shlobj.h>
#include <shlwapi.h>
#include <commctrl.h>
#include <cstdio>

Application& Application::Instance() {
    static Application inst;
    return inst;
}

bool Application::Init(HINSTANCE hInst, int nCmdShow) {
    m_hInst    = hInst;
    m_nCmdShow = nCmdShow;

    // Initialize Common Controls (visual styles, etc.)
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC  = ICC_BAR_CLASSES | ICC_COOL_CLASSES | ICC_STANDARD_CLASSES
               | ICC_TAB_CLASSES | ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    // Load persisted settings before creating the window
    LoadSettings();
    Localization::SetLanguage(m_settings.language);

    // Create the main window hidden. WM_APP+1 calls ShowWindow after the
    // editor and toolbar layout are fully ready, so the first visible paint
    // is already correct — the toolbar never needs a mouse hover to appear.
    m_hMainWnd = MainWindow::Create(hInst);
    if (!m_hMainWnd) return false;
    m_hAccel = LoadAcceleratorsW(m_hInst, MAKEINTRESOURCEW(IDR_ACCELERATORS));
    return true;
}

int Application::Run() {
    MSG msg{};
    while (GetMessage(&msg, nullptr, 0, 0)) {
        HWND hFind = MainWindow::GetFindDlg();
        if (hFind && IsWindow(hFind) && IsDialogMessage(hFind, &msg))
            continue;
        if (m_hAccel && TranslateAccelerator(m_hMainWnd, m_hAccel, &msg))
            continue;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    SaveSettings();
    return static_cast<int>(msg.wParam);
}

void Application::Quit() {
    PostQuitMessage(0);
}

std::wstring Application::SettingsPath() const {
    wchar_t path[MAX_PATH] = {};
    SHGetFolderPathW(nullptr, CSIDL_APPDATA, nullptr, SHGFP_TYPE_CURRENT, path);
    PathAppendW(path, L"WhatsUp");
    CreateDirectoryW(path, nullptr);
    PathAppendW(path, L"settings.ini");
    return path;
}

void Application::SaveSettings() {
    auto ini = SettingsPath();
    const wchar_t* sec = L"General";
    WritePrivateProfileStringW(sec, L"Language",
        m_settings.language == Language::Korean ? L"ko" : L"en", ini.c_str());
    WritePrivateProfileStringW(sec, L"Theme",
        m_settings.theme == Theme::Dark ? L"dark" : L"light", ini.c_str());
    WritePrivateProfileStringW(sec, L"UserName",   m_settings.userName.c_str(),   ini.c_str());
    WritePrivateProfileStringW(sec, L"UserOrg",    m_settings.userOrg.c_str(),    ini.c_str());
    WritePrivateProfileStringW(sec, L"SpellLang",  m_settings.spellLanguage.c_str(), ini.c_str());

    auto boolStr = [](bool v) { return v ? L"1" : L"0"; };
    WritePrivateProfileStringW(sec, L"Spell",        boolStr(m_settings.spellEnabled),  ini.c_str());
    WritePrivateProfileStringW(sec, L"Autocomplete", boolStr(m_settings.autocomplete),  ini.c_str());
    WritePrivateProfileStringW(sec, L"GridShow",     boolStr(m_settings.gridShow),      ini.c_str());
    WritePrivateProfileStringW(sec, L"GridSnap",     boolStr(m_settings.gridSnap),      ini.c_str());

    wchar_t buf[16];
    _snwprintf_s(buf, _countof(buf), _TRUNCATE, L"%d", m_settings.gridSize);
    WritePrivateProfileStringW(sec, L"GridSize", buf, ini.c_str());
}

void Application::LoadSettings() {
    auto ini = SettingsPath();
    const wchar_t* sec = L"General";
    wchar_t buf[256] = {};

    GetPrivateProfileStringW(sec, L"Language", L"ko", buf, _countof(buf), ini.c_str());
    m_settings.language = (wcscmp(buf, L"en") == 0) ? Language::English : Language::Korean;

    GetPrivateProfileStringW(sec, L"Theme", L"light", buf, _countof(buf), ini.c_str());
    m_settings.theme = (wcscmp(buf, L"dark") == 0) ? Theme::Dark : Theme::Light;

    GetPrivateProfileStringW(sec, L"UserName", L"", buf, _countof(buf), ini.c_str());
    m_settings.userName = buf;

    GetPrivateProfileStringW(sec, L"UserOrg", L"", buf, _countof(buf), ini.c_str());
    m_settings.userOrg = buf;

    GetPrivateProfileStringW(sec, L"SpellLang", L"ko-KR", buf, _countof(buf), ini.c_str());
    m_settings.spellLanguage = buf;

    auto boolVal = [&](const wchar_t* key, bool def) {
        return GetPrivateProfileIntW(sec, key, def ? 1 : 0, ini.c_str()) != 0;
    };
    m_settings.spellEnabled  = boolVal(L"Spell",        true);
    m_settings.autocomplete  = boolVal(L"Autocomplete", true);
    m_settings.gridShow      = boolVal(L"GridShow",     false);
    m_settings.gridSnap      = boolVal(L"GridSnap",     false);
    m_settings.gridSize      = GetPrivateProfileIntW(sec, L"GridSize", 10, ini.c_str());
}

COLORREF Application::TextColor() const {
    return (m_settings.theme == Theme::Dark) ? RGB(220, 220, 220) : RGB(0, 0, 0);
}

COLORREF Application::BgColor() const {
    return (m_settings.theme == Theme::Dark) ? RGB(30, 30, 30) : RGB(255, 255, 255);
}

COLORREF Application::ToolbarBgColor() const {
    return (m_settings.theme == Theme::Dark) ? RGB(45, 45, 48) : RGB(240, 240, 240);
}
