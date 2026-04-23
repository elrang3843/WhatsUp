#include "SplashScreen.h"
#include <cmath>
#include <cstdio>
#include <vector>
#include <algorithm>

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

int       SplashScreen::s_progress = 0;
HWND      SplashScreen::s_hwnd     = nullptr;
UINT_PTR  SplashScreen::s_timerId  = 0;

// ---- Main entry ----
void SplashScreen::Show(HINSTANCE hInst, int totalMs) {
    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClass;
    RegisterClassExW(&wc);

    // Center on primary monitor
    int sw = GetSystemMetrics(SM_CXSCREEN);
    int sh = GetSystemMetrics(SM_CYSCREEN);
    int w  = 900, h = 560;
    int x  = (sw - w) / 2, y = (sh - h) / 2;

    s_hwnd = CreateWindowExW(
        WS_EX_TOPMOST,
        kClass, L"What's Up",
        WS_POPUP,
        x, y, w, h,
        nullptr, nullptr, hInst, nullptr);

    if (!s_hwnd) return;
    ShowWindow(s_hwnd, SW_SHOW);
    UpdateWindow(s_hwnd);

    // Progress timer: fire every ~28ms for ~100 steps
    int stepMs = totalMs / 100;
    s_timerId = SetTimer(s_hwnd, 1, static_cast<UINT>(stepMs > 0 ? stepMs : 28), nullptr);

    // Pump until WM_DESTROY posts WM_QUIT; loop consumes it so the main
    // window's pump starts with a clean queue.
    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    UnregisterClassW(kClass, hInst);
}

// ---- Window procedure ----
LRESULT CALLBACK SplashScreen::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        OnPaint(hwnd, hdc);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_TIMER:
        if (wParam == 1) {
            s_progress = std::min(s_progress + 1, 100);
            InvalidateRect(hwnd, nullptr, FALSE);
            if (s_progress >= 100)
                DestroyWindow(hwnd);
        }
        return 0;
    case WM_DESTROY:
        KillTimer(hwnd, s_timerId);
        s_hwnd = nullptr;
        PostQuitMessage(0);
        return 0;
    case WM_LBUTTONDOWN:
        // Allow clicking to skip splash
        s_progress = 100;
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ---- Drawing ----
void SplashScreen::OnPaint(HWND hwnd, HDC hdc) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int w = rc.right, h = rc.bottom;

    // Double-buffer
    HDC memDC    = CreateCompatibleDC(hdc);
    HBITMAP hBmp = CreateCompatibleBitmap(hdc, w, h);
    SelectObject(memDC, hBmp);

    // Background
    DrawGradientBackground(memDC, rc);

    // cy is shifted higher so text rows + progress bar fit without overlap
    int cx = w / 2, cy = h / 2 - 80;   // 200 when h=560

    // Glow
    DrawGlowCircle(memDC, cx, cy, 120);

    // Gear logo
    DrawLogoIcon(memDC, cx, cy, 200);

    // "What's Up" title  [cy+120 .. cy+180]
    HFONT hTitleFont = CreateFontW(56, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                                   DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                   CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                   DEFAULT_PITCH, L"Segoe UI");
    RECT rcTitle{ 0, cy + 120, w, cy + 180 };
    DrawText_(memDC, L"What's Up", hTitleFont, RGB(255, 255, 255), rcTitle, DT_CENTER | DT_SINGLELINE);
    DeleteObject(hTitleFont);

    // "TEXT EDITOR" subtitle  [cy+184 .. cy+210]
    HFONT hSubFont = CreateFontW(22, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                  DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                  CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                  DEFAULT_PITCH, L"Segoe UI");
    RECT rcSub{ 0, cy + 184, w, cy + 210 };
    DrawText_(memDC, L"TEXT EDITOR", hSubFont, RGB(200, 220, 255), rcSub, DT_CENTER | DT_SINGLELINE);

    // "Open Source · Handtech"  [cy+214 .. cy+238]
    RECT rcBrand{ 0, cy + 214, w, cy + 238 };
    DrawText_(memDC, L"Open Source  ·  Handtech", hSubFont, RGB(130, 170, 230), rcBrand, DT_CENTER | DT_SINGLELINE);
    DeleteObject(hSubFont);

    // Progress bar
    static const wchar_t* const s_steps[] = {
        L"Loading modules...", L"Loading modules...",
        L"Initializing spell checker...", L"Initializing spell checker...",
        L"Loading fonts...",   L"Loading fonts...",
        L"Preparing editor...",L"Preparing editor...",
        L"Ready.",
    };
    int stepIdx = s_progress * 8 / 100;
    if (stepIdx > 8) stepIdx = 8;

    HFONT hSmall = CreateFontW(15, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                DEFAULT_PITCH, L"Segoe UI");
    RECT rcMsg{ 0, h - 90, w, h - 72 };
    DrawText_(memDC, s_steps[stepIdx], hSmall, RGB(200, 220, 255), rcMsg, DT_CENTER | DT_SINGLELINE);
    DeleteObject(hSmall);

    RECT rcBar{ w / 2 - 200, h - 66, w / 2 + 200, h - 50 };
    DrawProgressBar(memDC, rcBar, s_progress / 100.f);

    // Version
    HFONT hVer = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                              DEFAULT_PITCH, L"Segoe UI");
    RECT rcVer{ w - 90, h - 28, w - 10, h - 8 };
    DrawText_(memDC, L"v1.0.0", hVer, RGB(120, 150, 200), rcVer, DT_RIGHT | DT_SINGLELINE);
    DeleteObject(hVer);

    // Blit
    BitBlt(hdc, 0, 0, w, h, memDC, 0, 0, SRCCOPY);
    DeleteObject(hBmp);
    DeleteDC(memDC);
}

void SplashScreen::DrawGradientBackground(HDC hdc, const RECT& rc) {
    int w = rc.right - rc.left;
    int h = rc.bottom - rc.top;

    // Manual radial-ish gradient: dark blue at edges → slightly lighter center
    COLORREF outer = RGB(10, 25, 60);
    COLORREF inner = RGB(15, 45, 95);

    // Fill with horizontal gradient bands
    for (int y = 0; y < h; ++y) {
        float fy = static_cast<float>(y) / h; // 0 top, 1 bottom
        // Darken at top and bottom
        float fade = 1.0f - 0.35f * (fy < 0.5f ? (1.0f - 2*fy) : (2*fy - 1.0f));
        int r = static_cast<int>(GetRValue(inner) * fade + GetRValue(outer) * (1 - fade));
        int g = static_cast<int>(GetGValue(inner) * fade + GetGValue(outer) * (1 - fade));
        int b = static_cast<int>(GetBValue(inner) * fade + GetBValue(outer) * (1 - fade));
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(r, g, b));
        HPEN old  = reinterpret_cast<HPEN>(SelectObject(hdc, hPen));
        MoveToEx(hdc, 0, y, nullptr);
        LineTo(hdc, w, y);
        SelectObject(hdc, old);
        DeleteObject(hPen);
    }
}

void SplashScreen::DrawGlowCircle(HDC hdc, int cx, int cy, int r) {
    // Draw concentric semi-transparent rings to fake a glow
    for (int i = r; i > r - 40; --i) {
        int alpha = 255 - static_cast<int>((r - i) * 6.5f);
        if (alpha < 0) alpha = 0;
        HPEN hPen = CreatePen(PS_SOLID, 1, RGB(40, 100, 200));
        HBRUSH hBr = reinterpret_cast<HBRUSH>(GetStockObject(NULL_BRUSH));
        HPEN  oldP = reinterpret_cast<HPEN>(SelectObject(hdc, hPen));
        HBRUSH oldB = reinterpret_cast<HBRUSH>(SelectObject(hdc, hBr));
        Ellipse(hdc, cx - i, cy - i, cx + i, cy + i);
        SelectObject(hdc, oldP);
        SelectObject(hdc, oldB);
        DeleteObject(hPen);
    }
}

// Draw a gear shape using polygon
void SplashScreen::DrawGear(HDC hdc, int cx, int cy, int outerR, int innerR, int teeth) {
    const int pts = teeth * 4;
    std::vector<POINT> poly(pts);
    for (int i = 0; i < pts; ++i) {
        double angle = 2.0 * M_PI * i / pts;
        // Alternate between tooth tip and root
        bool tip = (i % 4 == 1 || i % 4 == 2);
        int  r   = tip ? outerR : innerR;
        // Slight offset for square teeth
        double a = angle + (i % 4 >= 2 ? M_PI / pts : 0);
        poly[i].x = cx + static_cast<int>(r * cos(a));
        poly[i].y = cy + static_cast<int>(r * sin(a));
    }
    Polygon(hdc, poly.data(), pts);
}

void SplashScreen::DrawLogoIcon(HDC hdc, int cx, int cy, int size) {
    int outerR  = size / 2;
    int innerR  = static_cast<int>(outerR * 0.82);

    // Gear body
    HBRUSH gearBrush = CreateSolidBrush(RGB(25, 55, 110));
    HPEN   gearPen   = CreatePen(PS_SOLID, 2, RGB(50, 100, 180));
    SelectObject(hdc, gearBrush);
    SelectObject(hdc, gearPen);
    DrawGear(hdc, cx, cy, outerR, innerR, 16);
    DeleteObject(gearBrush);
    DeleteObject(gearPen);

    // Inner circle (darker)
    HBRUSH innerBr = CreateSolidBrush(RGB(12, 30, 70));
    HPEN   innerPn = CreatePen(PS_SOLID, 1, RGB(40, 80, 160));
    SelectObject(hdc, innerBr);
    SelectObject(hdc, innerPn);
    int ir = static_cast<int>(outerR * 0.66);
    Ellipse(hdc, cx - ir, cy - ir, cx + ir, cy + ir);
    DeleteObject(innerBr);
    DeleteObject(innerPn);

    // "T" shape (Handtech logo silhouette) — two dark rounded rectangles
    HBRUSH tBrush = CreateSolidBrush(RGB(8, 20, 55));
    SelectObject(hdc, tBrush);
    SelectObject(hdc, GetStockObject(NULL_PEN));
    int tw = static_cast<int>(outerR * 0.50);
    int th = static_cast<int>(outerR * 0.08);
    int tv = static_cast<int>(outerR * 0.36);
    // Horizontal bar of T
    RECT rtH{ cx - tw/2, cy - th/2, cx + tw/2, cy + th/2 };
    FillRect(hdc, &rtH, tBrush);
    // Vertical bar of T
    RECT rtV{ cx - static_cast<int>(outerR*0.09), cy + th/2,
              cx + static_cast<int>(outerR*0.09), cy + th/2 + tv };
    FillRect(hdc, &rtV, tBrush);
    DeleteObject(tBrush);

    // Document icon (light blue square with lines)
    int docX = cx - static_cast<int>(outerR * 0.22);
    int docY = cy - static_cast<int>(outerR * 0.42);
    int docW = static_cast<int>(outerR * 0.38);
    int docH = static_cast<int>(outerR * 0.44);
    HBRUSH docBr = CreateSolidBrush(RGB(50, 100, 190));
    HPEN   docPn = CreatePen(PS_SOLID, 1, RGB(80, 150, 230));
    SelectObject(hdc, docBr); SelectObject(hdc, docPn);
    RoundRect(hdc, docX, docY, docX + docW, docY + docH, 6, 6);
    DeleteObject(docBr); DeleteObject(docPn);

    // Lines on document
    HPEN linePen = CreatePen(PS_SOLID, 2, RGB(160, 200, 255));
    SelectObject(hdc, linePen);
    int lx1 = docX + 5, lx2 = docX + docW - 5;
    for (int li = 0; li < 3; ++li) {
        int ly = docY + 8 + li * 9;
        MoveToEx(hdc, lx1, ly, nullptr);
        LineTo(hdc, lx2 - (li == 2 ? docW/3 : 0), ly);
    }
    DeleteObject(linePen);

    // Pencil icon (yellow-orange diagonal)
    HPEN pencilPen   = CreatePen(PS_SOLID, 5, RGB(240, 160, 30));
    HPEN pencilTip   = CreatePen(PS_SOLID, 3, RGB(200, 80, 20));
    SelectObject(hdc, pencilPen);
    int px1 = cx + static_cast<int>(outerR * 0.08);
    int py1 = cy - static_cast<int>(outerR * 0.05);
    int px2 = cx - static_cast<int>(outerR * 0.12);
    int py2 = cy + static_cast<int>(outerR * 0.28);
    MoveToEx(hdc, px1, py1, nullptr);
    LineTo(hdc, px2, py2);
    SelectObject(hdc, pencilTip);
    MoveToEx(hdc, px2, py2, nullptr);
    LineTo(hdc, px2 - 4, py2 + 6);
    DeleteObject(pencilPen);
    DeleteObject(pencilTip);
}

void SplashScreen::DrawProgressBar(HDC hdc, const RECT& rc, float pct) {
    // Track (dark rounded rect)
    HBRUSH trackBr = CreateSolidBrush(RGB(20, 40, 90));
    HPEN   trackPn = CreatePen(PS_SOLID, 1, RGB(40, 70, 140));
    SelectObject(hdc, trackBr);
    SelectObject(hdc, trackPn);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 8, 8);
    DeleteObject(trackBr); DeleteObject(trackPn);

    // Fill (gradient blue)
    int fillW = static_cast<int>((rc.right - rc.left) * pct);
    if (fillW > 2) {
        HBRUSH fillBr = CreateSolidBrush(RGB(60, 140, 255));
        HPEN   fillPn = CreatePen(PS_SOLID, 1, RGB(100, 180, 255));
        SelectObject(hdc, fillBr);
        SelectObject(hdc, fillPn);
        RECT fillRc{ rc.left, rc.top, rc.left + fillW, rc.bottom };
        RoundRect(hdc, fillRc.left, fillRc.top, fillRc.right, fillRc.bottom, 8, 8);
        DeleteObject(fillBr); DeleteObject(fillPn);
    }
}

void SplashScreen::DrawText_(HDC hdc, const wchar_t* text, HFONT hfont,
                              COLORREF color, const RECT& rc, UINT fmt) {
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, color);
    HFONT old = reinterpret_cast<HFONT>(SelectObject(hdc, hfont));
    RECT mrc = rc;
    DrawTextW(hdc, text, -1, &mrc, fmt);
    SelectObject(hdc, old);
}
