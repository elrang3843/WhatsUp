#pragma once
#include <windows.h>
#include <gdiplus.h>

class SplashScreen {
public:
    static void Show(HINSTANCE hInst, int totalMs = 2800);

private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void OnPaint(HWND hwnd, HDC hdc);

    static void DrawGradientBackground(HDC hdc, const RECT& rc);
    static void DrawProgressBar(HDC hdc, const RECT& rc, float pct);
    static void DrawText_(HDC hdc, const wchar_t* text, HFONT hfont,
                          COLORREF color, const RECT& rc, UINT fmt);

    static void LoadSplashImage();

    static int            s_progress;
    static HWND           s_hwnd;
    static UINT_PTR       s_timerId;
    static ULONG_PTR      s_gdiplusToken;
    static Gdiplus::Image* s_splashImage;

    static constexpr wchar_t kClass[] = L"WhatsUpSplash";
};
