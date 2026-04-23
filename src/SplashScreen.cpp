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

    // Document + pencil logo (gear drawn inside DrawLogoIcon)
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
    DrawText_(memDC, L"TEXT EDITOR", hSubFont, RGB(126, 184, 255), rcSub, DT_CENTER | DT_SINGLELINE);

    // Divider line (matches SVG: 136px wide, 3px tall, #4a90d9)
    {
        int dw = 68, dy = cy + 213;
        HBRUSH divBr = CreateSolidBrush(RGB(74, 144, 217));
        RECT divRc{ w/2 - dw, dy, w/2 + dw, dy + 3 };
        FillRect(memDC, &divRc, divBr);
        DeleteObject(divBr);
    }

    // "Open Source · Handtech"  [cy+218 .. cy+240]
    RECT rcBrand{ 0, cy + 218, w, cy + 240 };
    DrawText_(memDC, L"Open Source  \xB7  Handtech", hSubFont, RGB(74, 106, 154), rcBrand, DT_CENTER | DT_SINGLELINE);
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

// Gear polygon: teeth alternate between outerR (tip) and innerR (root).
void SplashScreen::DrawGear(HDC hdc, int cx, int cy, int outerR, int innerR, int teeth) {
    const int pts = teeth * 4;
    std::vector<POINT> poly(pts);
    for (int i = 0; i < pts; ++i) {
        double angle = 2.0 * M_PI * i / pts;
        bool   tip   = (i % 4 == 1 || i % 4 == 2);
        int    r     = tip ? outerR : innerR;
        double a     = angle + (i % 4 >= 2 ? M_PI / pts : 0.0);
        poly[i].x = cx + static_cast<int>(r * cos(a));
        poly[i].y = cy + static_cast<int>(r * sin(a));
    }
    Polygon(hdc, poly.data(), pts);
}

// Icon: large background gear → dark hub circle → document + pencil.
// Matches SVG structure: gear image at opacity 0.55, then document+pencil on top.
void SplashScreen::DrawLogoIcon(HDC hdc, int cx, int cy, int size) {
    float s = size / 200.0f;

    // ── Background gear (SVG: base64 image at opacity 0.55) ─────────────────
    // outerR ≈ 80 % of size; innerR ≈ 82 % of outer (shallow teeth)
    int gOuter = (int)(160.0f * s + 0.5f);
    int gInner = (int)(gOuter * 0.82f + 0.5f);

    // Color chosen to read "55 % opacity blue-gray over dark navy background"
    HBRUSH gearBr = CreateSolidBrush(RGB(22, 52, 105));
    HPEN   gearPn = CreatePen(PS_SOLID, std::max(1, (int)(2.0f*s+0.5f)), RGB(38, 78, 148));
    SelectObject(hdc, gearBr);
    SelectObject(hdc, gearPn);
    DrawGear(hdc, cx, cy, gOuter, gInner, 16);
    DeleteObject(gearBr);
    DeleteObject(gearPn);

    // Hub circle (slightly darker centre of the gear)
    int hubR = (int)(gOuter * 0.65f + 0.5f);
    HBRUSH hubBr = CreateSolidBrush(RGB(10, 26, 65));
    HPEN   hubPn = CreatePen(PS_SOLID, 1, RGB(22, 50, 100));
    SelectObject(hdc, hubBr);
    SelectObject(hdc, hubPn);
    Ellipse(hdc, cx - hubR, cy - hubR, cx + hubR, cy + hubR);
    DeleteObject(hubBr);
    DeleteObject(hubPn);

    // ── Document (rounded-rect outline, dark fill, 3 lines) ─────────────────
    // SVG coords relative to icon centre: x∈[-58,58], y∈[-65,33], rx=11
    int dl = cx - (int)(58*s + 0.5f), dr = cx + (int)(58*s + 0.5f);
    int dt = cy - (int)(65*s + 0.5f), db = cy + (int)(33*s + 0.5f);
    int drx = (int)(22*s + 0.5f);

    HBRUSH docBg  = CreateSolidBrush(RGB(8, 20, 50));
    HPEN   docPen = CreatePen(PS_SOLID, std::max(1, (int)(5*s + 0.5f)), RGB(106, 174, 255));
    SelectObject(hdc, docBg);
    SelectObject(hdc, docPen);
    RoundRect(hdc, dl, dt, dr, db, drx, drx);
    DeleteObject(docBg);
    DeleteObject(docPen);

    // Three horizontal lines
    int lw = std::max(1, (int)(4.5f*s + 0.5f));
    HPEN lp = CreatePen(PS_SOLID, lw, RGB(106, 174, 255));
    SelectObject(hdc, lp);
    int lx1 = cx - (int)(38*s + 0.5f), lx2 = cx + (int)(38*s + 0.5f);
    int lx3 = cx + (int)(10*s + 0.5f);
    int lys[3] = { cy - (int)(40*s + 0.5f),
                   cy - (int)(18*s + 0.5f),
                   cy + (int)( 5*s + 0.5f) };
    MoveToEx(hdc, lx1, lys[0], nullptr); LineTo(hdc, lx2, lys[0]);
    MoveToEx(hdc, lx1, lys[1], nullptr); LineTo(hdc, lx2, lys[1]);
    MoveToEx(hdc, lx1, lys[2], nullptr); LineTo(hdc, lx3, lys[2]);
    DeleteObject(lp);

    // ── Pencil (rotated -42°, translated to icon-relative (30,28)) ──────────
    float px = cx + 30.0f * s, py = cy + 28.0f * s;
    float ca = 0.7431f, sa = -0.6691f;   // cos/sin of -42°

    auto rot = [&](float lx, float ry) -> POINT {
        return { (LONG)(lx * ca - ry * sa + px + 0.5f),
                 (LONG)(lx * sa + ry * ca + py + 0.5f) };
    };

    float pw  =  7.0f * s;
    float pbt = -32.0f * s, pbb = 18.0f * s, ptt = 34.0f * s;
    float pet = -41.0f * s, peb = -31.0f * s;

    HPEN noPen = (HPEN)GetStockObject(NULL_PEN);
    SelectObject(hdc, noPen);

    { POINT p[4] = { rot(-pw,pet), rot(pw,pet), rot(pw,peb), rot(-pw,peb) };
      HBRUSH br = CreateSolidBrush(RGB(192, 192, 200));
      SelectObject(hdc, br); Polygon(hdc, p, 4); DeleteObject(br); }

    { POINT p[4] = { rot(-pw,pbt), rot(pw,pbt), rot(pw,pbb), rot(-pw,pbb) };
      HBRUSH br = CreateSolidBrush(RGB(255, 209, 102));
      SelectObject(hdc, br); Polygon(hdc, p, 4); DeleteObject(br); }

    { POINT p[3] = { rot(-pw,pbb), rot(pw,pbb), rot(0,ptt) };
      HBRUSH br = CreateSolidBrush(RGB(240, 160, 32));
      SelectObject(hdc, br); Polygon(hdc, p, 3); DeleteObject(br); }
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
