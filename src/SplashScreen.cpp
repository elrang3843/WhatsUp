#include "SplashScreen.h"
#include "resource.h"
#include <cmath>
#include <algorithm>

#ifndef M_PI
#define M_PI 3.14159265358979323846
#endif

int             SplashScreen::s_progress     = 0;
HWND            SplashScreen::s_hwnd         = nullptr;
UINT_PTR        SplashScreen::s_timerId      = 0;
ULONG_PTR       SplashScreen::s_gdiplusToken = 0;
Gdiplus::Image* SplashScreen::s_splashImage  = nullptr;

// ---- Load PNG: embedded resource first, then file-system fallback ----------
void SplashScreen::LoadSplashImage() {
    // ── Try embedded RCDATA resource ─────────────────────────────────────────
    // NOTE: RCDATA keyword in .rc stores as numeric type RT_RCDATA (10).
    //       Passing the string L"RCDATA" would look for a custom named type and
    //       always return NULL. Must use the RT_RCDATA macro.
    HMODULE hMod  = GetModuleHandleW(nullptr);
    HRSRC   hRsrc = FindResourceW(hMod, MAKEINTRESOURCEW(IDR_SPLASH_PNG), RT_RCDATA);
    if (hRsrc) {
        HGLOBAL hGlob = LoadResource(hMod, hRsrc);
        DWORD   sz    = SizeofResource(hMod, hRsrc);
        void*   pSrc  = LockResource(hGlob);
        if (pSrc && sz > 0) {
            HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, sz);
            if (hMem) {
                void* pDst = GlobalLock(hMem);
                if (pDst) {
                    CopyMemory(pDst, pSrc, sz);
                    GlobalUnlock(hMem);
                    IStream* pStream = nullptr;
                    if (SUCCEEDED(CreateStreamOnHGlobal(hMem, TRUE, &pStream))) {
                        s_splashImage = Gdiplus::Image::FromStream(pStream);
                        pStream->Release();
                        if (s_splashImage &&
                            s_splashImage->GetLastStatus() != Gdiplus::Ok) {
                            delete s_splashImage;
                            s_splashImage = nullptr;
                        }
                    } else {
                        GlobalFree(hMem);
                    }
                } else {
                    GlobalFree(hMem);
                }
            }
        }
    }

    // ── Fallback: load from file system (e.g. resource not embedded yet) ─────
    if (!s_splashImage) {
        static const wchar_t* kPaths[] = {
            L"images\\WhatsUp_start1.png",
            L"..\\images\\WhatsUp_start1.png",
            L"WhatsUp_start1.png",
        };
        for (auto p : kPaths) {
            auto* img = Gdiplus::Image::FromFile(p);
            if (img && img->GetLastStatus() == Gdiplus::Ok) {
                s_splashImage = img;
                break;
            }
            delete img;
        }
    }
}

// ---- Main entry ------------------------------------------------------------
void SplashScreen::Show(HINSTANCE hInst, int totalMs) {
    // Initialise GDI+ and load image before creating the window
    Gdiplus::GdiplusStartupInput gsi;
    Gdiplus::GdiplusStartup(&s_gdiplusToken, &gsi, nullptr);
    LoadSplashImage();

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kClass;
    RegisterClassExW(&wc);

    int sw = GetSystemMetrics(SM_CXSCREEN);
    int sh = GetSystemMetrics(SM_CYSCREEN);
    int w  = 900, h = 560;
    int x  = (sw - w) / 2, y = (sh - h) / 2;

    s_hwnd = CreateWindowExW(
        WS_EX_TOPMOST, kClass, L"What's Up", WS_POPUP,
        x, y, w, h, nullptr, nullptr, hInst, nullptr);

    if (!s_hwnd) {
        delete s_splashImage; s_splashImage = nullptr;
        Gdiplus::GdiplusShutdown(s_gdiplusToken);
        return;
    }
    ShowWindow(s_hwnd, SW_SHOW);
    UpdateWindow(s_hwnd);

    int stepMs = totalMs / 100;
    s_timerId  = SetTimer(s_hwnd, 1, static_cast<UINT>(stepMs > 0 ? stepMs : 28), nullptr);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    delete s_splashImage; s_splashImage = nullptr;
    Gdiplus::GdiplusShutdown(s_gdiplusToken);
    UnregisterClassW(kClass, hInst);
}

// ---- Window procedure ------------------------------------------------------
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
        s_progress = 100;
        DestroyWindow(hwnd);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ---- Drawing ---------------------------------------------------------------
void SplashScreen::OnPaint(HWND hwnd, HDC hdc) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    int w = rc.right, h = rc.bottom;

    // Double-buffer
    HDC     memDC = CreateCompatibleDC(hdc);
    HBITMAP hBmp  = CreateCompatibleBitmap(hdc, w, h);
    SelectObject(memDC, hBmp);

    DrawGradientBackground(memDC, rc);

    // ── PNG splash image ───────────────────────────────────────────────────
    // The PNG already contains the icon, gear, and all text labels.
    // Scale it to fill the upper portion of the splash; leave ~110px at the
    // bottom for the loading message and progress bar.
    if (s_splashImage) {
        UINT imgW = s_splashImage->GetWidth();
        UINT imgH = s_splashImage->GetHeight();

        int availH = h - 110;       // reserve bottom for progress bar area
        int availW = w - 40;        // 20px margin each side

        float scaleH = static_cast<float>(availH) / imgH;
        float scaleW = static_cast<float>(availW) / imgW;
        float scale  = std::min(scaleH, scaleW);

        int destW = static_cast<int>(imgW * scale);
        int destH = static_cast<int>(imgH * scale);
        int destX = (w - destW) / 2;
        int destY = std::max(0, (availH - destH) / 2);

        Gdiplus::Graphics g(memDC);
        g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
        g.SetSmoothingMode(Gdiplus::SmoothingModeAntiAlias);
        g.DrawImage(s_splashImage,
                    Gdiplus::Rect(destX, destY, destW, destH));
    }

    // ── Loading message ────────────────────────────────────────────────────
    static const wchar_t* const s_steps[] = {
        L"Loading modules...",        L"Loading modules...",
        L"Initializing spell checker...", L"Initializing spell checker...",
        L"Loading fonts...",          L"Loading fonts...",
        L"Preparing editor...",       L"Preparing editor...",
        L"Ready.",
    };
    int stepIdx = s_progress * 8 / 100;
    if (stepIdx > 8) stepIdx = 8;

    HFONT hSmall = CreateFontW(15, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                                CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                                DEFAULT_PITCH, L"Segoe UI");
    RECT rcMsg{ 0, h - 90, w, h - 72 };
    DrawText_(memDC, s_steps[stepIdx], hSmall, RGB(200, 220, 255), rcMsg,
              DT_CENTER | DT_SINGLELINE);
    DeleteObject(hSmall);

    // ── Progress bar ───────────────────────────────────────────────────────
    RECT rcBar{ w / 2 - 200, h - 66, w / 2 + 200, h - 50 };
    DrawProgressBar(memDC, rcBar, s_progress / 100.f);

    // ── Version label (bottom-right) ───────────────────────────────────────
    HFONT hVer = CreateFontW(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                              DEFAULT_PITCH, L"Segoe UI");
    RECT rcVer{ w - 90, h - 28, w - 10, h - 8 };
    DrawText_(memDC, L"v1.0.0", hVer, RGB(120, 150, 200), rcVer,
              DT_RIGHT | DT_SINGLELINE);
    DeleteObject(hVer);

    BitBlt(hdc, 0, 0, w, h, memDC, 0, 0, SRCCOPY);
    DeleteObject(hBmp);
    DeleteDC(memDC);
}

// ---- Helpers ---------------------------------------------------------------
void SplashScreen::DrawGradientBackground(HDC hdc, const RECT& rc) {
    int w = rc.right  - rc.left;
    int h = rc.bottom - rc.top;

    COLORREF outer = RGB(10, 25, 60);
    COLORREF inner = RGB(15, 45, 95);

    for (int y = 0; y < h; ++y) {
        float fy   = static_cast<float>(y) / h;
        float fade = 1.0f - 0.35f * (fy < 0.5f ? (1.0f - 2*fy) : (2*fy - 1.0f));
        int   r    = static_cast<int>(GetRValue(inner)*fade + GetRValue(outer)*(1-fade));
        int   g2   = static_cast<int>(GetGValue(inner)*fade + GetGValue(outer)*(1-fade));
        int   b    = static_cast<int>(GetBValue(inner)*fade + GetBValue(outer)*(1-fade));
        HPEN  hPen = CreatePen(PS_SOLID, 1, RGB(r, g2, b));
        HPEN  old  = reinterpret_cast<HPEN>(SelectObject(hdc, hPen));
        MoveToEx(hdc, 0, y, nullptr);
        LineTo(hdc, w, y);
        SelectObject(hdc, old);
        DeleteObject(hPen);
    }
}

void SplashScreen::DrawProgressBar(HDC hdc, const RECT& rc, float pct) {
    HBRUSH trackBr = CreateSolidBrush(RGB(20, 40, 90));
    HPEN   trackPn = CreatePen(PS_SOLID, 1, RGB(40, 70, 140));
    SelectObject(hdc, trackBr); SelectObject(hdc, trackPn);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 8, 8);
    DeleteObject(trackBr); DeleteObject(trackPn);

    int fillW = static_cast<int>((rc.right - rc.left) * pct);
    if (fillW > 2) {
        HBRUSH fillBr = CreateSolidBrush(RGB(60, 140, 255));
        HPEN   fillPn = CreatePen(PS_SOLID, 1, RGB(100, 180, 255));
        SelectObject(hdc, fillBr); SelectObject(hdc, fillPn);
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
    RECT  mrc = rc;
    DrawTextW(hdc, text, -1, &mrc, fmt);
    SelectObject(hdc, old);
}
