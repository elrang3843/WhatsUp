#pragma once
#include <windows.h>
#include <richedit.h>
#include <string>
#include <vector>
#include <functional>

#ifndef CLR_NONE
#define CLR_NONE ((COLORREF)-1)
#endif

struct CharFormat {
    std::wstring fontName;
    int          fontSize    = 0;   // half-points (0 = unchanged)
    COLORREF     textColor   = CLR_NONE;
    COLORREF     bgColor     = CLR_NONE;
    int          bold        = -1;  // -1 = unchanged, 0 = off, 1 = on
    int          italic      = -1;
    int          underline   = -1;
    int          strikeout   = -1;
    int          superscript = -1;
    int          subscript   = -1;
    int          outline     = -1;
    int          textWidth   = 0;   // percentage scaling (0 = unchanged)
    int          spacing     = 0;   // twips (0 = unchanged)
};

struct ParaFormat {
    int  alignment    = -1;  // PFA_LEFT/CENTER/RIGHT/JUSTIFY; -1 = unchanged
    int  leftIndent   = -1;  // twips
    int  rightIndent  = -1;
    int  firstIndent  = 0;
    int  spaceBefore  = -1;  // twips
    int  spaceAfter   = -1;
    int  lineSpacing  = 0;   // twips; 0 = single
    int  numbering    = -1;  // 0 = none, PFN_BULLET, PFN_ARABIC, etc.
};

struct TextStats {
    int words      = 0;
    int chars      = 0;
    int charNoSpace= 0;
    int lines      = 0;
    int paragraphs = 0;
};

class Editor {
public:
    Editor();
    ~Editor();

    bool Create(HWND hwndParent, UINT controlId);
    void Destroy();

    HWND GetHwnd() const { return m_hwnd; }

    // Resize to fill rect
    void Move(int x, int y, int cx, int cy);

    // ---- Content ----
    std::wstring GetText() const;
    void         SetText(const std::wstring& text);
    std::string  GetRtf() const;
    void         SetRtf(const std::string& rtf);
    void         Clear();

    // ---- Selection ----
    void GetSel(int& start, int& end) const;
    void SetSel(int start, int end);
    std::wstring GetSelText() const;
    void ReplaceSelection(const std::wstring& text, bool canUndo = true);

    // ---- Caret info ----
    int  GetCaretPos() const;
    int  GetCaretLine() const;
    int  GetCaretCol() const;
    int  GetLineCount() const;

    // ---- Edit operations ----
    void Undo()       { SendMessageW(m_hwnd, WM_UNDO, 0, 0); }
    void Redo()       { SendMessageW(m_hwnd, EM_REDO, 0, 0); }
    void Cut()        { SendMessageW(m_hwnd, WM_CUT, 0, 0); }
    void Copy()       { SendMessageW(m_hwnd, WM_COPY, 0, 0); }
    void Paste()      { SendMessageW(m_hwnd, WM_PASTE, 0, 0); }
    void SelectAll()  { SendMessageW(m_hwnd, EM_SETSEL, 0, -1); }

    bool CanUndo() const { return m_hwnd && SendMessageW(m_hwnd, EM_CANUNDO, 0, 0) != 0; }
    bool CanRedo() const { return m_hwnd && SendMessageW(m_hwnd, EM_CANREDO, 0, 0) != 0; }
    bool CanPaste()const { return m_hwnd && SendMessageW(m_hwnd, EM_CANPASTE, 0, 0) != 0; }

    // ---- Character formatting ----
    void ApplyCharFormat(const CharFormat& cf, bool selOnly = true);
    CharFormat GetCharFormat(bool selOnly = true) const;

    // Quick formatting shortcuts
    void ToggleBold();
    void ToggleItalic();
    void ToggleUnderline();
    void ToggleStrikeout();
    void ToggleSuperscript();
    void ToggleSubscript();
    void SetFontColor(COLORREF color);
    void SetBgColor(COLORREF color);
    void SetFont(const std::wstring& name, int halfPoints);

    // ---- Paragraph formatting ----
    void ApplyParaFormat(const ParaFormat& pf, bool selOnly = true);
    ParaFormat GetParaFormat() const;

    void SetAlignment(WORD alignment); // PFA_LEFT etc.
    void Indent(int direction);        // +1 = indent, -1 = outdent (twips = 360)
    void ToggleBullet();
    void ToggleNumbering();

    // ---- Find / Replace ----
    struct FindOptions {
        std::wstring findText;
        std::wstring replaceText;
        bool matchCase  = false;
        bool wholeWord  = false;
        bool forward    = true;
    };

    bool FindNext(const FindOptions& opts);
    int  ReplaceAll(const FindOptions& opts);

    // ---- Statistics ----
    TextStats ComputeStats() const;

    // ---- Spell check underline ----
    void MarkMisspelled(int start, int length, bool mark);
    void ClearAllMisspellings();

    // ---- Print ----
    bool Print(HWND hwndOwner, const std::wstring& docTitle);

    // ---- Modified flag ----
    bool IsModified() const { return m_hwnd && SendMessageW(m_hwnd, EM_GETMODIFY, 0, 0) != 0; }
    void SetModified(bool v){ if (m_hwnd) SendMessageW(m_hwnd, EM_SETMODIFY, v ? TRUE : FALSE, 0); }

    // ---- Background color (for theme) ----
    void SetBackground(COLORREF color);

    // ---- Scroll position ----
    void SetScrollPos(int y);

    // ---- Insert ----
    void InsertText(const std::wstring& text);
    void InsertHyperlink(const std::wstring& displayText, const std::wstring& url);
    void InsertTable(int rows, int cols, bool header, int widthPct);

private:
    HWND    m_hwnd     = nullptr;
    HMODULE m_hRichEd  = nullptr;

    // RTF stream callbacks
    static DWORD CALLBACK RtfReadCallback(DWORD_PTR cookie, LPBYTE buf, LONG size, LONG* pBytes);
    static DWORD CALLBACK RtfWriteCallback(DWORD_PTR cookie, LPBYTE buf, LONG size, LONG* pBytes);
};
