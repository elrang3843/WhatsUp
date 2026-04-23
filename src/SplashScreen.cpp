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
    DrawGlowCircle(memDC, cx, cy, 90);

    // Document + pencil logo
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

// Document + pencil icon matching the SVG design.
// SVG reference: icon center (256,190) in 512×512; we scale to 'size' pixels.
void SplashScreen::DrawLogoIcon(HDC hdc, int cx, int cy, int size) {
    float s = size / 200.0f;

    // ── Document (rounded-rect outline, dark fill, 3 lines) ──────────────────
    // SVG coords relative to icon centre: x∈[-58,58], y∈[-65,33], rx=11
    int dl = cx - (int)(58*s + 0.5f), dr = cx + (int)(58*s + 0.5f);
    int dt = cy - (int)(65*s + 0.5f), db = cy + (int)(33*s + 0.5f);
    int drx = (int)(22*s + 0.5f);  // RoundRect takes diameter

    HBRUSH docBg  = CreateSolidBrush(RGB(8, 20, 50));
    HPEN   docPen = CreatePen(PS_SOLID, std::max(1, (int)(5*s + 0.5f)), RGB(106, 174, 255));
    SelectObject(hdc, docBg);
    SelectObject(hdc, docPen);
    RoundRect(hdc, dl, dt, dr, db, drx, drx);
    DeleteObject(docBg);
    DeleteObject(docPen);

    // Three horizontal lines (SVG: opacity 0.75, stroke-width 4.5)
    int lw  = std::max(1, (int)(4.5f*s + 0.5f));
    HPEN lp = CreatePen(PS_SOLID, lw, RGB(106, 174, 255));
    SelectObject(hdc, lp);
    int lx1 = cx - (int)(38*s + 0.5f), lx2 = cx + (int)(38*s + 0.5f);
    int lx3 = cx + (int)(10*s + 0.5f);          // line 3 is shorter
    int lys[3] = { cy - (int)(40*s + 0.5f),
                   cy - (int)(18*s + 0.5f),
                   cy + (int)( 5*s + 0.5f) };
    MoveToEx(hdc, lx1, lys[0], nullptr); LineTo(hdc, lx2, lys[0]);
    MoveToEx(hdc, lx1, lys[1], nullptr); LineTo(hdc, lx2, lys[1]);
    MoveToEx(hdc, lx1, lys[2], nullptr); LineTo(hdc, lx3, lys[2]);
    DeleteObject(lp);

    // ── Pencil (rotated -42°, translated to icon-relative (30,28)) ───────────
    // cos(-42°) ≈ 0.7431, sin(-42°) ≈ -0.6691
    float px  = cx + 30.0f * s, py  = cy + 28.0f * s;
    float ca  = 0.7431f,        sa  = -0.6691f;

    // Rotate local pencil point (lx,ry) and map to screen
    auto rot = [&](float lx, float ry) -> POINT {
        return { (LONG)(lx * ca - ry * sa + px + 0.5f),
                 (LONG)(lx * sa + ry * ca + py + 0.5f) };
    };

    float pw  =  7.0f * s;   // half-width of pencil shaft
    float pbt = -32.0f * s;  // body top
    float pbb =  18.0f * s;  // body bottom  (= tip base)
    float ptt =  34.0f * s;  // tip apex
    float pet = -41.0f * s;  // eraser top
    float peb = -31.0f * s;  // eraser bottom (= body top - 1px gap)

    HPEN noPen = (HPEN)GetStockObject(NULL_PEN);
    SelectObject(hdc, noPen);

    // Eraser cap (light gray)
    { POINT p[4] = { rot(-pw, pet), rot(pw, pet), rot(pw, peb), rot(-pw, peb) };
      HBRUSH br = CreateSolidBrush(RGB(192, 192, 200));
      SelectObject(hdc, br); Polygon(hdc, p, 4); DeleteObject(br); }

    // Pencil body (golden yellow)
    { POINT p[4] = { rot(-pw, pbt), rot(pw, pbt), rot(pw, pbb), rot(-pw, pbb) };
      HBRUSH br = CreateSolidBrush(RGB(255, 209, 102));
      SelectObject(hdc, br); Polygon(hdc, p, 4); DeleteObject(br); }

    // Pencil tip (orange triangle)
    { POINT p[3] = { rot(-pw, pbb), rot(pw, pbb), rot(0, ptt) };
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
