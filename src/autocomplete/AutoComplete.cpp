#include "AutoComplete.h"
#include <algorithm>
#include <sstream>
#include <richedit.h>

static constexpr int AC_LIST_MAX_VISIBLE = 8;
static constexpr int AC_LIST_ITEM_HEIGHT = 18;

AutoComplete::AutoComplete() = default;
AutoComplete::~AutoComplete() { Destroy(); }

bool AutoComplete::Create(HWND hwndParent, HWND hwndEditor) {
    m_hwndParent = hwndParent;
    m_hwndEditor = hwndEditor;

    // Create a popup listbox (no parent WS_POPUP so it floats)
    m_hwndList = CreateWindowExW(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        L"LISTBOX", nullptr,
        WS_POPUP | WS_BORDER | LBS_NOTIFY | LBS_HASSTRINGS | WS_VSCROLL,
        0, 0, 200, AC_LIST_ITEM_HEIGHT * AC_LIST_MAX_VISIBLE,
        hwndParent, nullptr, GetModuleHandleW(nullptr), nullptr);

    if (!m_hwndList) return false;

    // Set font to match editor
    HFONT hFont = reinterpret_cast<HFONT>(SendMessageW(hwndEditor, WM_GETFONT, 0, 0));
    if (hFont) SendMessageW(m_hwndList, WM_SETFONT, reinterpret_cast<WPARAM>(hFont), TRUE);

    // Subclass the listbox to intercept key input
    SetWindowSubclass(m_hwndList, ListProc, 1, reinterpret_cast<DWORD_PTR>(this));
    return true;
}

void AutoComplete::Destroy() {
    HidePopup();
    if (m_hwndList) { DestroyWindow(m_hwndList); m_hwndList = nullptr; }
}

void AutoComplete::OnTextChanged(const std::wstring& fullText, int caretPos) {
    std::wstring prefix = GetPrefix(fullText, caretPos);
    if (prefix.size() < static_cast<size_t>(m_minPrefixLen)) {
        HidePopup();
        return;
    }
    if (prefix == m_currentPrefix && IsVisible()) return;

    m_currentPrefix  = prefix;
    m_currentMatches = FindMatches(prefix);

    if (m_currentMatches.empty()) {
        HidePopup();
        return;
    }
    ShowPopup(m_hwndEditor, prefix, m_currentMatches);
}

bool AutoComplete::HandleKey(UINT vk) {
    if (!IsVisible()) return false;

    switch (vk) {
    case VK_ESCAPE:
        HidePopup();
        return true;
    case VK_RETURN:
    case VK_TAB:
        AcceptSelection();
        return true;
    case VK_DOWN: {
        int cur = static_cast<int>(SendMessageW(m_hwndList, LB_GETCURSEL, 0, 0));
        int cnt = static_cast<int>(SendMessageW(m_hwndList, LB_GETCOUNT, 0, 0));
        if (cur + 1 < cnt) SendMessageW(m_hwndList, LB_SETCURSEL, cur + 1, 0);
        return true;
    }
    case VK_UP: {
        int cur = static_cast<int>(SendMessageW(m_hwndList, LB_GETCURSEL, 0, 0));
        if (cur > 0) SendMessageW(m_hwndList, LB_SETCURSEL, cur - 1, 0);
        return true;
    }
    default:
        return false;
    }
}

void AutoComplete::AcceptSelection() {
    if (!m_hwndList) return;
    int idx = static_cast<int>(SendMessageW(m_hwndList, LB_GETCURSEL, 0, 0));
    if (idx < 0 || idx >= static_cast<int>(m_currentMatches.size())) {
        HidePopup();
        return;
    }
    const std::wstring& word = m_currentMatches[idx];
    // Replace the prefix with the full word in the editor
    std::wstring suffix = word.substr(m_currentPrefix.size());
    SendMessageW(m_hwndEditor, EM_REPLACESEL, TRUE,
                 reinterpret_cast<LPARAM>(suffix.c_str()));
    HidePopup();
}

void AutoComplete::ShowPopup(HWND hwndEditor, const std::wstring& /*prefix*/,
                              const std::vector<std::wstring>& matches) {
    SendMessageW(m_hwndList, LB_RESETCONTENT, 0, 0);
    for (const auto& w : matches)
        SendMessageW(m_hwndList, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(w.c_str()));
    SendMessageW(m_hwndList, LB_SETCURSEL, 0, 0);

    PositionPopup(hwndEditor, -1);
    int h = std::min(static_cast<int>(matches.size()), AC_LIST_MAX_VISIBLE) * AC_LIST_ITEM_HEIGHT;
    RECT rc;
    GetWindowRect(m_hwndList, &rc);
    SetWindowPos(m_hwndList, HWND_TOPMOST, rc.left, rc.top,
                 rc.right - rc.left, h, SWP_NOMOVE | SWP_SHOWWINDOW);
}

void AutoComplete::HidePopup() {
    if (m_hwndList) ShowWindow(m_hwndList, SW_HIDE);
    m_currentPrefix.clear();
    m_currentMatches.clear();
}

void AutoComplete::PositionPopup(HWND hwndEditor, int /*caretPos*/) {
    // Position the popup below the caret in the editor
    POINT caretPt{};
    DWORD selStart = 0, selEnd = 0;
    SendMessageW(hwndEditor, EM_GETSEL,
                 reinterpret_cast<WPARAM>(&selStart),
                 reinterpret_cast<LPARAM>(&selEnd));
    POINTL pt{};
    SendMessageW(hwndEditor, EM_POSFROMCHAR,
                 reinterpret_cast<WPARAM>(&pt),
                 static_cast<LPARAM>(selEnd));
    caretPt = { pt.x, pt.y };
    ClientToScreen(hwndEditor, &caretPt);

    RECT rc;
    GetWindowRect(m_hwndList, &rc);
    int w = rc.right - rc.left;
    SetWindowPos(m_hwndList, HWND_TOPMOST,
                 caretPt.x, caretPt.y + 18,
                 w, rc.bottom - rc.top,
                 SWP_NOSIZE | SWP_NOZORDER);
}

std::wstring AutoComplete::GetPrefix(const std::wstring& text, int caretPos) {
    if (caretPos <= 0 || caretPos > static_cast<int>(text.size())) return {};
    int i = caretPos - 1;
    while (i >= 0 && (iswalpha(text[i]) || text[i] == L'_' || iswdigit(text[i])))
        --i;
    return text.substr(i + 1, caretPos - i - 1);
}

std::vector<std::wstring> AutoComplete::FindMatches(const std::wstring& prefix,
                                                     int maxResults) const {
    std::vector<std::wstring> matches;
    std::wstring lower = prefix;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::towlower);

    for (const auto& word : m_dict) {
        if (word.size() <= prefix.size()) continue;
        std::wstring wl = word;
        std::transform(wl.begin(), wl.end(), wl.begin(), ::towlower);
        if (wl.substr(0, lower.size()) == lower)
            matches.push_back(word);
        if (static_cast<int>(matches.size()) >= maxResults) break;
    }
    std::sort(matches.begin(), matches.end());
    return matches;
}

void AutoComplete::LearnText(const std::wstring& text) {
    std::wstring word;
    for (wchar_t ch : text) {
        if (iswalpha(ch) || ch == L'_' || (iswdigit(ch) && !word.empty())) {
            word += ch;
        } else {
            if (word.size() >= 2) m_dict.insert(word);
            word.clear();
        }
    }
    if (word.size() >= 2) m_dict.insert(word);
}

void AutoComplete::ClearDictionary() {
    m_dict.clear();
}

LRESULT CALLBACK AutoComplete::ListProc(HWND hwnd, UINT msg, WPARAM wParam,
                                         LPARAM lParam, UINT_PTR /*id*/,
                                         DWORD_PTR data) {
    auto* ac = reinterpret_cast<AutoComplete*>(data);
    if (msg == WM_LBUTTONDBLCLK) {
        ac->AcceptSelection();
        return 0;
    }
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}
