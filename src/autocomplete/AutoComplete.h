#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <unordered_set>

// Autocomplete popup: monitors the editor for word-prefix typed,
// shows a floating LISTBOX with completions from a learned dictionary.
class AutoComplete {
public:
    AutoComplete();
    ~AutoComplete();

    bool Create(HWND hwndParent, HWND hwndEditor);
    void Destroy();

    // Call after each character typed in the editor
    void OnTextChanged(const std::wstring& fullText, int caretPos);

    // Returns true if autocomplete popup is visible
    bool IsVisible() const { return m_hwndList && IsWindowVisible(m_hwndList); }

    // Handle keyboard in editor to navigate/accept suggestions
    // Returns true if the key was consumed by autocomplete
    bool HandleKey(UINT vk);

    // Learn words from a document
    void LearnText(const std::wstring& text);
    void ClearDictionary();

    // Minimum prefix length to trigger autocomplete
    void SetMinPrefixLength(int len) { m_minPrefixLen = len; }

private:
    void  ShowPopup(HWND hwndEditor, const std::wstring& prefix,
                    const std::vector<std::wstring>& matches);
    void  HidePopup();
    void  AcceptSelection();
    void  PositionPopup(HWND hwndEditor, int caretPos);

    // Extract the current word prefix at caretPos
    static std::wstring GetPrefix(const std::wstring& text, int caretPos);

    // Find all words starting with prefix (up to maxResults)
    std::vector<std::wstring> FindMatches(const std::wstring& prefix, int maxResults = 10) const;

    HWND    m_hwndParent = nullptr;
    HWND    m_hwndEditor = nullptr;
    HWND    m_hwndList   = nullptr;
    int     m_minPrefixLen = 2;

    std::wstring m_currentPrefix;
    std::vector<std::wstring> m_currentMatches;

    // Dictionary: all unique words seen in documents
    std::unordered_set<std::wstring> m_dict;

    static LRESULT CALLBACK ListProc(HWND hwnd, UINT msg, WPARAM wParam,
                                      LPARAM lParam, UINT_PTR id, DWORD_PTR data);
};
