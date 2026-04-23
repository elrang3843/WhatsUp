#pragma once
#include <windows.h>

// Startup splash screen matching the design:
// dark blue radial gradient, gear logo (GDI), progress bar, version text.
class SplashScreen {
public:
    // Show the splash, run the progress animation, then destroy.
    // totalMs: approximate total display time in milliseconds.
    static void Show(HINSTANCE hInst, int totalMs = 2800);

private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    static void OnPaint(HWND hwnd, HDC hdc);

    // GDI drawing helpers
    static void DrawGradientBackground(HDC hdc, const RECT& rc);
    static void DrawGlowCircle(HDC hdc, int cx, int cy, int r);
    static void DrawGear(HDC hdc, int cx, int cy, int outerR, int innerR, int teeth);
    static void DrawLogoIcon(HDC hdc, int cx, int cy, int size);
    static void DrawProgressBar(HDC hdc, const RECT& rc, float pct);
    static void DrawText_(HDC hdc, const wchar_t* text, HFONT hfont,
                          COLORREF color, const RECT& rc, UINT fmt);

    static int   s_progress;   // 0–100
    static HWND  s_hwnd;

    static constexpr wchar_t kClass[] = L"WhatsUpSplash";
};
