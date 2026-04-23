#pragma once
#include <windows.h>
#include <string>
#include "i18n/Localization.h"

enum class Theme { Light, Dark };

struct AppSettings {
    Language    language        = Language::Korean;
    Theme       theme           = Theme::Light;
    std::wstring userName;
    std::wstring userOrg;
    bool        spellEnabled    = true;
    std::wstring spellLanguage  = L"ko-KR";
    bool        autocomplete    = true;
    bool        gridShow        = false;
    bool        gridSnap        = false;
    int         gridSize        = 10;     // pixels
};

class Application {
public:
    static Application& Instance();

    bool    Init(HINSTANCE hInst, int nCmdShow);
    int     Run();
    void    Quit();

    HINSTANCE   GetHInstance() const { return m_hInst; }
    HWND        GetMainWindow() const { return m_hMainWnd; }
    int         GetNCmdShow()  const { return m_nCmdShow; }

    AppSettings& Settings() { return m_settings; }
    const AppSettings& Settings() const { return m_settings; }

    void    SaveSettings();
    void    LoadSettings();

    // Theme helpers
    COLORREF TextColor() const;
    COLORREF BgColor() const;
    COLORREF ToolbarBgColor() const;

private:
    Application() = default;
    Application(const Application&) = delete;
    Application& operator=(const Application&) = delete;

    HINSTANCE   m_hInst    = nullptr;
    HWND        m_hMainWnd = nullptr;
    HACCEL      m_hAccel   = nullptr;
    int         m_nCmdShow = SW_SHOWNORMAL;
    AppSettings m_settings;

    std::wstring SettingsPath() const;
};
