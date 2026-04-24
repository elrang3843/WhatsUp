#include "Editor.h"
#include <sstream>
#include <algorithm>
#include <commdlg.h>
#include <winspool.h>
#include <imm.h>
#include <objbase.h>      // IStream (prereq for gdiplus.h)
#include <gdiplus.h>
#include <olectl.h>       // IPicture, OleCreatePictureIndirect, PICTDESC

#ifndef CFM_UNDERLINECOLOR
#define CFM_UNDERLINECOLOR 0x00800000
#endif

Editor::Editor() {
    Gdiplus::GdiplusStartupInput gsi;
    Gdiplus::GdiplusStartup(&m_gdiplusToken, &gsi, nullptr);
}
Editor::~Editor() {
    Destroy();
    if (m_gdiplusToken) {
        Gdiplus::GdiplusShutdown(m_gdiplusToken);
        m_gdiplusToken = 0;
    }
}

// Write a one-line entry to %TEMP%\whatsup_richedit.log for diagnosis.
static void RichEditLog(const wchar_t* msg) {
    wchar_t path[MAX_PATH];
    GetTempPathW(MAX_PATH, path);
    wcscat_s(path, L"whatsup_richedit.log");
    FILE* f = nullptr;
    _wfopen_s(&f, path, L"a");
    if (f) { fwprintf(f, L"%s\n", msg); fclose(f); }
    OutputDebugStringW(msg);
    OutputDebugStringW(L"\n");
}

bool Editor::Create(HWND hwndParent, UINT controlId) {
    RichEditLog(L"[WhatsUp] Editor::Create() called");

    // Candidate DLL + class name pairs.  On Windows 10+ riched20.dll is a
    // thin stub that loads msftedit.dll, so both entries ultimately use the
    // same underlying DLL.  We try Msftedit (newer) first, then Riched20.
    struct { LPCWSTR dll; LPCWSTR cls; } kEdits[] = {
        { L"Msftedit.dll", MSFTEDIT_CLASS },   // RICHEDIT50W
        { L"Riched20.dll", RICHEDIT_CLASS  },   // RichEdit20W
    };

    for (auto& e : kEdits) {
        wchar_t dbg[128];
        swprintf_s(dbg, L"[WhatsUp] Trying %s / %s", e.dll, e.cls);
        RichEditLog(dbg);

        HMODULE h = LoadLibraryW(e.dll);
        if (!h) {
            swprintf_s(dbg, L"[WhatsUp]   LoadLibrary FAILED err=%lu", GetLastError());
            RichEditLog(dbg);
            continue;
        }
        RichEditLog(L"[WhatsUp]   LoadLibrary OK");

        // msftedit.dll spawns a D2D/DirectWrite init thread.  Retry up to
        // 5 times (250 ms total) to give that thread time to complete.
        HWND hw = nullptr;
        for (int retry = 0; retry < 5 && !hw; ++retry) {
            if (retry > 0) Sleep(50);
            hw = CreateWindowExW(
                WS_EX_CLIENTEDGE, e.cls, nullptr,
                WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL
                | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL
                | ES_NOHIDESEL | ES_WANTRETURN,
                0, 0, 0, 0,
                hwndParent,
                reinterpret_cast<HMENU>(static_cast<UINT_PTR>(controlId)),
                GetModuleHandleW(nullptr), nullptr);
            swprintf_s(dbg, L"[WhatsUp]   CreateWindowExW try %d: %s (err=%lu)",
                       retry+1, hw ? L"OK" : L"FAIL", GetLastError());
            RichEditLog(dbg);
        }

        if (hw) {
            m_hRichEd = h;
            m_hwnd    = hw;
            break;
        }
        FreeLibrary(h);
    }

    if (!m_hwnd) {
        RichEditLog(L"[WhatsUp] All RichEdit attempts failed");
        wchar_t buf[512];
        swprintf_s(buf, _countof(buf),
            L"RichEdit 컨트롤을 생성할 수 없습니다.\n"
            L"Msftedit.dll 및 Riched20.dll 모두 실패.\n\n"
            L"로그 파일: %%TEMP%%\\whatsup_richedit.log\n"
            L"마지막 오류: %lu (0x%08lX)",
            GetLastError(), GetLastError());
        MessageBoxW(hwndParent, buf, L"WhatsUp 초기화 오류", MB_OK | MB_ICONERROR);
        return false;
    }
    RichEditLog(L"[WhatsUp] RichEdit window created OK");

    // Rich-text mode so EM_SETCHARFORMAT with SCF_ALL/SCF_DEFAULT works correctly
    // after WM_SETTEXT. TM_PLAINTEXT suppresses those formatting calls, making
    // text invisible in dark theme after a file load.
    SendMessageW(m_hwnd, EM_SETTEXTMODE, TM_RICHTEXT | TM_MULTILEVELUNDO, 0);
    SendMessageW(m_hwnd, EM_SETUNDOLIMIT, 100, 0);
    SendMessageW(m_hwnd, EM_SETEVENTMASK,
                 ENM_CHANGE | ENM_SELCHANGE | ENM_SCROLL | ENM_KEYEVENTS, 0);

    // Enable IME (Korean input)
    ImmAssociateContextEx(m_hwnd, nullptr, IACE_DEFAULT);

    // Set default font (맑은 고딕 or Malgun Gothic)
    CharFormat cf{};
    cf.fontName = L"맑은 고딕";
    cf.fontSize = 22; // 11pt in half-points
    ApplyCharFormat(cf, false);

    // Word wrap
    SendMessageW(m_hwnd, EM_SETTARGETDEVICE, 0, 0); // auto wrap

    return true;
}

void Editor::Destroy() {
    if (m_hwnd) { DestroyWindow(m_hwnd); m_hwnd = nullptr; }
    if (m_hRichEd) { FreeLibrary(m_hRichEd); m_hRichEd = nullptr; }
}

void Editor::Move(int x, int y, int cx, int cy) {
    if (m_hwnd) MoveWindow(m_hwnd, x, y, cx, cy, TRUE);
}

// ---- Content ----

std::wstring Editor::GetText() const {
    int len = static_cast<int>(SendMessageW(m_hwnd, WM_GETTEXTLENGTH, 0, 0));
    if (len <= 0) return {};
    std::wstring text(len + 1, 0);
    SendMessageW(m_hwnd, WM_GETTEXT, len + 1, reinterpret_cast<LPARAM>(text.data()));
    text.resize(len);
    return text;
}

void Editor::SetText(const std::wstring& text) {
    SendMessageW(m_hwnd, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(text.c_str()));
    SetModified(false);
}

// RTF streaming
struct RtfReadCtx  { const std::string* data; size_t pos; };
struct RtfWriteCtx { std::string* data; };

DWORD CALLBACK Editor::RtfReadCallback(DWORD_PTR cookie, LPBYTE buf, LONG size, LONG* pBytes) {
    auto* ctx = reinterpret_cast<RtfReadCtx*>(cookie);
    LONG  avail = static_cast<LONG>(ctx->data->size() - ctx->pos);
    *pBytes = std::min(size, avail);
    memcpy(buf, ctx->data->data() + ctx->pos, *pBytes);
    ctx->pos += *pBytes;
    return 0;
}

DWORD CALLBACK Editor::RtfWriteCallback(DWORD_PTR cookie, LPBYTE buf, LONG size, LONG* pBytes) {
    auto* ctx = reinterpret_cast<RtfWriteCtx*>(cookie);
    ctx->data->append(reinterpret_cast<char*>(buf), size);
    *pBytes = size;
    return 0;
}

std::string Editor::GetRtf() const {
    std::string rtf;
    RtfWriteCtx ctx{ &rtf };
    EDITSTREAM es{};
    es.dwCookie    = reinterpret_cast<DWORD_PTR>(&ctx);
    es.pfnCallback = RtfWriteCallback;
    SendMessageW(m_hwnd, EM_STREAMOUT, SF_RTF, reinterpret_cast<LPARAM>(&es));
    return rtf;
}

void Editor::SetRtf(const std::string& rtf) {
    {
        wchar_t dbg[256];
        size_t  preview = (rtf.size() < 120) ? rtf.size() : 120;
        wchar_t head[128] = {};
        for (size_t i = 0; i < preview; ++i)
            head[i] = static_cast<wchar_t>(static_cast<unsigned char>(rtf[i]));
        swprintf_s(dbg, L"[WhatsUp] SetRtf %zu bytes: %s", rtf.size(), head);
        RichEditLog(dbg);
    }

    RtfReadCtx ctx{ &rtf, 0 };
    EDITSTREAM es{};
    es.dwCookie    = reinterpret_cast<DWORD_PTR>(&ctx);
    es.pfnCallback = RtfReadCallback;
    LRESULT res = SendMessageW(m_hwnd, EM_STREAMIN, SF_RTF, reinterpret_cast<LPARAM>(&es));
    {
        wchar_t dbg[128];
        swprintf_s(dbg, L"[WhatsUp] EM_STREAMIN result=%lld dwError=%lu consumed=%zu",
                   static_cast<long long>(res), es.dwError, ctx.pos);
        RichEditLog(dbg);
    }

    // On error, dump the full RTF to a file for inspection
    if (es.dwError != 0) {
        wchar_t dumpPath[MAX_PATH];
        GetTempPathW(MAX_PATH, dumpPath);
        wcscat_s(dumpPath, L"whatsup_rtf_dump.txt");
        FILE* f = nullptr;
        _wfopen_s(&f, dumpPath, L"wb");
        if (f) {
            fwrite(rtf.data(), 1, rtf.size(), f);
            fclose(f);
            wchar_t dbg[MAX_PATH + 64];
            swprintf_s(dbg, L"[WhatsUp] RTF dump written to: %s", dumpPath);
            RichEditLog(dbg);
        }
    }

    SetModified(false);
}

void Editor::Clear() {
    SendMessageW(m_hwnd, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(L""));
    SetModified(false);
}

// ---- Selection ----

void Editor::GetSel(int& start, int& end) const {
    CHARRANGE cr{};
    SendMessageW(m_hwnd, EM_EXGETSEL, 0, reinterpret_cast<LPARAM>(&cr));
    start = cr.cpMin; end = cr.cpMax;
}

void Editor::SetSel(int start, int end) {
    CHARRANGE cr{ start, end };
    SendMessageW(m_hwnd, EM_EXSETSEL, 0, reinterpret_cast<LPARAM>(&cr));
}

std::wstring Editor::GetSelText() const {
    int s, e; GetSel(s, e);
    if (s == e) return {};
    int len = e - s;
    std::wstring text(len + 1, 0);
    SendMessageW(m_hwnd, EM_GETSELTEXT, 0, reinterpret_cast<LPARAM>(text.data()));
    text.resize(len);
    return text;
}

void Editor::ReplaceSelection(const std::wstring& text, bool canUndo) {
    SendMessageW(m_hwnd, EM_REPLACESEL, canUndo ? TRUE : FALSE,
                 reinterpret_cast<LPARAM>(text.c_str()));
}

// ---- Caret info ----

int Editor::GetCaretPos() const {
    int s, e; GetSel(s, e); return e;
}

int Editor::GetCaretLine() const {
    int pos = GetCaretPos();
    return static_cast<int>(SendMessageW(m_hwnd, EM_LINEFROMCHAR, pos, 0)) + 1;
}

int Editor::GetCaretCol() const {
    int pos = GetCaretPos();
    int line = static_cast<int>(SendMessageW(m_hwnd, EM_LINEFROMCHAR, pos, 0));
    int lineStart = static_cast<int>(SendMessageW(m_hwnd, EM_LINEINDEX, line, 0));
    return pos - lineStart + 1;
}

int Editor::GetLineCount() const {
    return static_cast<int>(SendMessageW(m_hwnd, EM_GETLINECOUNT, 0, 0));
}

// ---- Character formatting ----

void Editor::ApplyCharFormat(const CharFormat& cf, bool selOnly) {
    CHARFORMAT2W fmt{};
    fmt.cbSize = sizeof(fmt);

    if (!cf.fontName.empty()) {
        fmt.dwMask |= CFM_FACE;
        wcscpy_s(fmt.szFaceName, cf.fontName.c_str());
    }
    if (cf.fontSize > 0) {
        fmt.dwMask |= CFM_SIZE;
        fmt.yHeight = cf.fontSize * 5; // half-points → twips (1pt = 20 twips, 1 half-pt = 10 twips)
        // Actually: yHeight is in twips. 1pt = 20 twips. fontSize is half-points → pt = fontSize/2.
        // So twips = (fontSize / 2) * 20 = fontSize * 10.
        fmt.yHeight = cf.fontSize * 10;
    }
    if (cf.textColor != CLR_NONE) {
        fmt.dwMask |= CFM_COLOR;
        fmt.crTextColor = cf.textColor;
    }
    if (cf.bgColor != CLR_NONE) {
        fmt.dwMask |= CFM_BACKCOLOR;
        fmt.crBackColor = cf.bgColor;
    }

    auto setBit = [&](int val, DWORD mask, DWORD effect) {
        if (val < 0) return;
        fmt.dwMask    |= mask;
        fmt.dwEffects |= (val ? effect : 0);
        if (!val) fmt.dwEffects &= ~effect;
    };
    setBit(cf.bold,        CFM_BOLD,        CFE_BOLD);
    setBit(cf.italic,      CFM_ITALIC,      CFE_ITALIC);
    setBit(cf.underline,   CFM_UNDERLINE,   CFE_UNDERLINE);
    setBit(cf.strikeout,   CFM_STRIKEOUT,   CFE_STRIKEOUT);
    setBit(cf.superscript, CFM_SUPERSCRIPT, CFE_SUPERSCRIPT);
    setBit(cf.subscript,   CFM_SUBSCRIPT,   CFE_SUBSCRIPT);

    WPARAM scx = selOnly ? SCF_SELECTION : SCF_ALL;
    SendMessageW(m_hwnd, EM_SETCHARFORMAT, scx, reinterpret_cast<LPARAM>(&fmt));
    if (!selOnly)
        SendMessageW(m_hwnd, EM_SETCHARFORMAT, SCF_DEFAULT, reinterpret_cast<LPARAM>(&fmt));
}

CharFormat Editor::GetCharFormat(bool selOnly) const {
    CHARFORMAT2W fmt{};
    fmt.cbSize = sizeof(fmt);
    WPARAM scx = selOnly ? SCF_SELECTION : SCF_DEFAULT;
    SendMessageW(m_hwnd, EM_GETCHARFORMAT, scx, reinterpret_cast<LPARAM>(&fmt));

    CharFormat cf{};
    cf.fontName = fmt.szFaceName;
    cf.fontSize = static_cast<int>(fmt.yHeight / 10); // twips → half-points
    cf.textColor = (fmt.dwMask & CFM_COLOR)    ? fmt.crTextColor : CLR_NONE;
    cf.bgColor   = (fmt.dwMask & CFM_BACKCOLOR)? fmt.crBackColor  : CLR_NONE;
    cf.bold        = (fmt.dwEffects & CFE_BOLD)       ? 1 : 0;
    cf.italic      = (fmt.dwEffects & CFE_ITALIC)     ? 1 : 0;
    cf.underline   = (fmt.dwEffects & CFE_UNDERLINE)  ? 1 : 0;
    cf.strikeout   = (fmt.dwEffects & CFE_STRIKEOUT)  ? 1 : 0;
    cf.superscript = (fmt.dwEffects & CFE_SUPERSCRIPT)? 1 : 0;
    cf.subscript   = (fmt.dwEffects & CFE_SUBSCRIPT)  ? 1 : 0;
    return cf;
}

void Editor::ToggleBold() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.bold = cur.bold ? 0 : 1;
    ApplyCharFormat(cf, true);
}
void Editor::ToggleItalic() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.italic = cur.italic ? 0 : 1;
    ApplyCharFormat(cf, true);
}
void Editor::ToggleUnderline() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.underline = cur.underline ? 0 : 1;
    ApplyCharFormat(cf, true);
}
void Editor::ToggleStrikeout() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.strikeout = cur.strikeout ? 0 : 1;
    ApplyCharFormat(cf, true);
}
void Editor::ToggleSuperscript() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.superscript = cur.superscript ? 0 : 1;
    if (cf.superscript) cf.subscript = 0;
    ApplyCharFormat(cf, true);
}
void Editor::ToggleSubscript() {
    CharFormat cur = GetCharFormat(true);
    CharFormat cf{};
    cf.subscript = cur.subscript ? 0 : 1;
    if (cf.subscript) cf.superscript = 0;
    ApplyCharFormat(cf, true);
}
void Editor::SetFontColor(COLORREF color) {
    CharFormat cf{};
    cf.textColor = color;
    ApplyCharFormat(cf, true);
}
void Editor::SetBgColor(COLORREF color) {
    CharFormat cf{};
    cf.bgColor = color;
    ApplyCharFormat(cf, true);
}
void Editor::SetFont(const std::wstring& name, int halfPoints) {
    CharFormat cf{};
    cf.fontName = name;
    cf.fontSize = halfPoints;
    ApplyCharFormat(cf, false);
}

// ---- Paragraph formatting ----

void Editor::ApplyParaFormat(const ParaFormat& pf, bool /*selOnly*/) {
    PARAFORMAT2 fmt{};
    fmt.cbSize = sizeof(fmt);

    if (pf.alignment >= 0) {
        fmt.dwMask  |= PFM_ALIGNMENT;
        fmt.wAlignment = static_cast<WORD>(pf.alignment);
    }
    if (pf.leftIndent >= 0) {
        fmt.dwMask    |= PFM_STARTINDENT;
        fmt.dxStartIndent = pf.leftIndent;
    }
    if (pf.rightIndent >= 0) {
        fmt.dwMask    |= PFM_RIGHTINDENT;
        fmt.dxRightIndent = pf.rightIndent;
    }
    if (pf.lineSpacing != 0) {
        fmt.dwMask    |= PFM_LINESPACING;
        fmt.dyLineSpacing = pf.lineSpacing;
        fmt.bLineSpacingRule = 4; // single (0=none applied)
    }
    if (pf.numbering >= 0) {
        fmt.dwMask    |= PFM_NUMBERING;
        fmt.wNumbering = static_cast<WORD>(pf.numbering);
    }
    SendMessageW(m_hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
}

ParaFormat Editor::GetParaFormat() const {
    PARAFORMAT2 fmt{};
    fmt.cbSize = sizeof(fmt);
    SendMessageW(m_hwnd, EM_GETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
    ParaFormat pf{};
    pf.alignment    = fmt.wAlignment;
    pf.leftIndent   = fmt.dxStartIndent;
    pf.rightIndent  = fmt.dxRightIndent;
    pf.numbering    = fmt.wNumbering;
    return pf;
}

void Editor::SetAlignment(WORD alignment) {
    PARAFORMAT2 fmt{};
    fmt.cbSize     = sizeof(fmt);
    fmt.dwMask     = PFM_ALIGNMENT;
    fmt.wAlignment = alignment;
    SendMessageW(m_hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
}

void Editor::Indent(int direction) {
    PARAFORMAT2 fmt{};
    fmt.cbSize = sizeof(fmt);
    SendMessageW(m_hwnd, EM_GETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
    fmt.dwMask = PFM_STARTINDENT;
    fmt.dxStartIndent = std::max(0L, fmt.dxStartIndent + static_cast<LONG>(direction * 720));
    SendMessageW(m_hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
}

void Editor::ToggleBullet() {
    PARAFORMAT2 fmt{};
    fmt.cbSize = sizeof(fmt);
    SendMessageW(m_hwnd, EM_GETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
    fmt.dwMask     = PFM_NUMBERING | PFM_NUMBERINGSTART | PFM_OFFSET;
    fmt.wNumbering = (fmt.wNumbering == PFN_BULLET) ? 0 : PFN_BULLET;
    fmt.dxOffset   = (fmt.wNumbering != 0) ? 360 : 0;
    SendMessageW(m_hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
}

void Editor::ToggleNumbering() {
    PARAFORMAT2 fmt{};
    fmt.cbSize = sizeof(fmt);
    SendMessageW(m_hwnd, EM_GETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
    fmt.dwMask          = PFM_NUMBERING | PFM_NUMBERINGSTART | PFM_OFFSET;
    fmt.wNumbering      = (fmt.wNumbering == PFN_ARABIC) ? 0 : PFN_ARABIC;
    fmt.wNumberingStart = 1;
    fmt.dxOffset        = (fmt.wNumbering != 0) ? 360 : 0;
    SendMessageW(m_hwnd, EM_SETPARAFORMAT, 0, reinterpret_cast<LPARAM>(&fmt));
}

// ---- Find / Replace ----

bool Editor::FindNext(const FindOptions& opts) {
    if (opts.findText.empty()) return false;
    int s, e; GetSel(s, e);
    int searchFrom = opts.forward ? e : s - 1;

    FINDTEXTEXW ft{};
    if (opts.forward) {
        ft.chrg.cpMin = searchFrom;
        ft.chrg.cpMax = -1;
    } else {
        ft.chrg.cpMin = 0;
        ft.chrg.cpMax = searchFrom;
    }
    ft.lpstrText = opts.findText.c_str();

    WPARAM flags = 0;
    if (opts.forward)   flags |= FR_DOWN;
    if (opts.matchCase) flags |= FR_MATCHCASE;
    if (opts.wholeWord) flags |= FR_WHOLEWORD;

    int found = static_cast<int>(
        SendMessageW(m_hwnd, EM_FINDTEXTEXW, flags, reinterpret_cast<LPARAM>(&ft)));
    if (found < 0) return false;

    SetSel(ft.chrgText.cpMin, ft.chrgText.cpMax);
    SendMessageW(m_hwnd, EM_SCROLLCARET, 0, 0);
    return true;
}

int Editor::ReplaceAll(const FindOptions& opts) {
    if (opts.findText.empty()) return 0;
    int count = 0;
    SetSel(0, 0);
    FindOptions fwd = opts;
    fwd.forward = true;
    while (FindNext(fwd)) {
        ReplaceSelection(opts.replaceText, true);
        ++count;
    }
    return count;
}

// ---- Statistics ----

TextStats Editor::ComputeStats() const {
    std::wstring text = GetText();
    TextStats st{};
    st.chars = static_cast<int>(text.size());
    bool inWord = false;
    for (wchar_t ch : text) {
        if (ch == L'\n') { ++st.lines; ++st.paragraphs; }
        if (iswspace(ch)) {
            if (inWord) { ++st.words; inWord = false; }
        } else {
            inWord = true;
            ++st.charNoSpace;
        }
    }
    if (inWord) ++st.words;
    st.lines      = GetLineCount();
    st.paragraphs = std::max(1, st.paragraphs);
    return st;
}

// ---- Spell check underline ----

void Editor::MarkMisspelled(int start, int length, bool mark) {
    CHARFORMAT2W cf{};
    cf.cbSize        = sizeof(cf);
    cf.dwMask        = CFM_UNDERLINETYPE;
    cf.bUnderlineType= mark ? CFU_UNDERLINEWAVE : CFU_UNDERLINENONE;
    if (mark) {
        cf.dwMask    |= CFM_UNDERLINECOLOR;
        cf.bUnderlineColor = 0x02; // red wave
    }
    SetSel(start, start + length);
    SendMessageW(m_hwnd, EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&cf));
}

void Editor::ClearAllMisspellings() {
    int s, e; GetSel(s, e);
    CHARFORMAT2W cf{};
    cf.cbSize        = sizeof(cf);
    cf.dwMask        = CFM_UNDERLINETYPE;
    cf.bUnderlineType= CFU_UNDERLINENONE;
    SetSel(0, -1);
    SendMessageW(m_hwnd, EM_SETCHARFORMAT, SCF_SELECTION, reinterpret_cast<LPARAM>(&cf));
    SetSel(s, e);
}

// ---- Print ----

bool Editor::Print(HWND hwndOwner, const std::wstring& docTitle) {
    PRINTDLGW pd{};
    pd.lStructSize = sizeof(pd);
    pd.hwndOwner   = hwndOwner;
    pd.Flags       = PD_RETURNDC | PD_NOSELECTION | PD_ALLPAGES;
    pd.nFromPage   = 1; pd.nToPage = 1; pd.nMinPage = 1; pd.nMaxPage = 9999;

    if (!PrintDlgW(&pd)) return false;

    HDC hdc = pd.hDC;
    DOCINFOW di{};
    di.cbSize      = sizeof(di);
    di.lpszDocName = docTitle.c_str();
    StartDocW(hdc, &di);
    StartPage(hdc);

    // Format range to printer DC
    FORMATRANGE fr{};
    fr.hdc     = hdc;
    fr.hdcTarget = hdc;
    POINT physSize{ GetDeviceCaps(hdc, PHYSICALWIDTH), GetDeviceCaps(hdc, PHYSICALHEIGHT) };
    POINT physOff { GetDeviceCaps(hdc, PHYSICALOFFSETX), GetDeviceCaps(hdc, PHYSICALOFFSETY) };
    // Convert to twips (1 inch = 1440 twips)
    int   dpix = GetDeviceCaps(hdc, LOGPIXELSX);
    int   dpiy = GetDeviceCaps(hdc, LOGPIXELSY);
    fr.rcPage.left   = 0;
    fr.rcPage.top    = 0;
    fr.rcPage.right  = MulDiv(physSize.x, 1440, dpix);
    fr.rcPage.bottom = MulDiv(physSize.y, 1440, dpiy);
    fr.rc = fr.rcPage;
    InflateRect(&fr.rc, -MulDiv(physOff.x, 1440, dpix) - 720,
                        -MulDiv(physOff.y, 1440, dpiy) - 720);
    fr.chrg.cpMin = 0;
    fr.chrg.cpMax = -1;

    SendMessageW(m_hwnd, EM_FORMATRANGE, TRUE, reinterpret_cast<LPARAM>(&fr));
    SendMessageW(m_hwnd, EM_FORMATRANGE, FALSE, 0); // free cache

    EndPage(hdc);
    EndDoc(hdc);
    DeleteDC(hdc);
    if (pd.hDevMode)  GlobalFree(pd.hDevMode);
    if (pd.hDevNames) GlobalFree(pd.hDevNames);
    return true;
}

// ---- Background ----

void Editor::SetBackground(COLORREF color) {
    SendMessageW(m_hwnd, EM_SETBKGNDCOLOR, 0, color);
}

void Editor::SetScrollPos(int y) {
    SendMessageW(m_hwnd, EM_LINESCROLL, 0, y);
}

// ---- Insert ----

void Editor::InsertText(const std::wstring& text) {
    ReplaceSelection(text, true);
}

void Editor::InsertHyperlink(const std::wstring& displayText, const std::wstring& url) {
    // Insert as RTF with hyperlink field
    std::string rtf = "{\\rtf1\\ansi{\\field{\\*\\fldinst{HYPERLINK \"";
    // Escape URL
    for (char c : std::string(url.begin(), url.end()))
        rtf += c;
    rtf += "\"}}{\\fldrslt ";
    // Display text (encode as Unicode)
    rtf += "{\\ul ";
    for (wchar_t ch : displayText) {
        if (ch < 128) rtf += static_cast<char>(ch);
        else { rtf += "\\u"; rtf += std::to_string(static_cast<int>(ch)); rtf += "?"; }
    }
    rtf += "}}}}\r\n}";

    RtfReadCtx ctx{ &rtf, 0 };
    EDITSTREAM es{};
    es.dwCookie    = reinterpret_cast<DWORD_PTR>(&ctx);
    es.pfnCallback = RtfReadCallback;
    SendMessageW(m_hwnd, EM_STREAMIN, SF_RTF | SFF_SELECTION, reinterpret_cast<LPARAM>(&es));
}

void Editor::InsertTable(int rows, int cols, bool header, int widthPct) {
    // Rich Edit doesn't natively support tables; insert a text-art representation
    std::wostringstream tbl;
    // Header separator
    auto makeLine = [&]() {
        tbl << L"+";
        for (int c = 0; c < cols; ++c) tbl << L"-------------------+";
        tbl << L"\r\n";
    };
    makeLine();
    for (int r = 0; r < rows; ++r) {
        tbl << L"|";
        for (int c = 0; c < cols; ++c)
            tbl << L"                   |";
        tbl << L"\r\n";
        if (header && r == 0) makeLine();
        makeLine();
    }
    (void)widthPct;
    InsertText(tbl.str());
}

// Wrap a Gdiplus::Bitmap as an OLE IPicture. The resulting IPicture owns
// the underlying HBITMAP (fOwn=TRUE); caller must Release() when done.
// Returns nullptr on failure.
static IPicture* CreatePictureFromBitmap(Gdiplus::Bitmap& bmp) {
    HBITMAP hbmp = nullptr;
    if (bmp.GetHBITMAP(Gdiplus::Color(0, 0, 0, 0), &hbmp) != Gdiplus::Ok || !hbmp)
        return nullptr;

    PICTDESC pd{};
    pd.cbSizeofstruct = sizeof(pd);
    pd.picType        = PICTYPE_BITMAP;
    pd.bmp.hbitmap    = hbmp;
    pd.bmp.hpal       = nullptr;

    IPicture* pic = nullptr;
    HRESULT hr = OleCreatePictureIndirect(&pd, IID_IPicture, TRUE,
                                          reinterpret_cast<void**>(&pic));
    if (FAILED(hr) || !pic) {
        DeleteObject(hbmp);
        return nullptr;
    }
    return pic;
}

bool Editor::InsertImageAt(int start, int end,
                           const std::vector<uint8_t>& imageBytes,
                           int widthPx, int heightPx) {
    if (!m_hwnd || imageBytes.empty()) return false;

    // Decode the bytes through GDI+ to validate the image and to learn its
    // intrinsic size. B5c will feed the resulting pixels to RichEdit.
    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, imageBytes.size());
    if (!hGlobal) return false;
    void* p = GlobalLock(hGlobal);
    if (!p) { GlobalFree(hGlobal); return false; }
    memcpy(p, imageBytes.data(), imageBytes.size());
    GlobalUnlock(hGlobal);

    IStream* stream = nullptr;
    if (CreateStreamOnHGlobal(hGlobal, TRUE, &stream) != S_OK) {
        GlobalFree(hGlobal);
        return false;
    }

    Gdiplus::Bitmap bmp(stream);
    bool ok = (bmp.GetLastStatus() == Gdiplus::Ok);
    UINT intrinsicW = ok ? bmp.GetWidth()  : 0;
    UINT intrinsicH = ok ? bmp.GetHeight() : 0;

    stream->Release();  // TRUE above hands ownership of hGlobal to the stream.

    if (!ok) {
        RichEditLog(L"[WhatsUp] InsertImageAt: GDI+ decode failed");
        (void)start; (void)end; (void)widthPx; (void)heightPx;
        return false;
    }

    IPicture* picture = CreatePictureFromBitmap(bmp);
    if (!picture) {
        RichEditLog(L"[WhatsUp] InsertImageAt: OleCreatePictureIndirect failed");
        return false;
    }

    wchar_t buf[160];
    swprintf_s(buf, L"[WhatsUp] InsertImageAt: IPicture ready %ux%u (bytes=%zu, range=%d..%d, hint=%dx%d)",
               intrinsicW, intrinsicH, imageBytes.size(), start, end, widthPx, heightPx);
    RichEditLog(buf);

    // B5c-2/B5c-3 will attach the IPicture to an IOleClientSite and insert
    // it into the RichEdit control via IRichEditOle::InsertObject.
    picture->Release();
    return false;
}
