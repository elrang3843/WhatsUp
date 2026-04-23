#pragma once
#include <windows.h>
#include <string>
#include <vector>
#include <memory>

struct SpellSuggestion {
    std::wstring word;
};

// Uses Windows 8+ ISpellCheckerFactory COM interface.
// Falls back gracefully on older systems.
class SpellChecker {
public:
    SpellChecker();
    ~SpellChecker();

    bool    IsAvailable() const { return m_available; }
    bool    SetLanguage(const std::wstring& langTag); // e.g. "ko-KR", "en-US"

    // Check a single word; returns true if correct
    bool    Check(const std::wstring& word);

    // Get suggestions for a misspelled word
    std::vector<SpellSuggestion> Suggest(const std::wstring& word);

    // Find misspellings in a text range; returns start positions
    struct Mistake {
        int start;
        int length;
    };
    std::vector<Mistake> Check(const std::wstring& text, int startOffset = 0);

    // Add to user dictionary
    void    AddWord(const std::wstring& word);
    void    IgnoreWord(const std::wstring& word);

private:
    void    Init();

    bool    m_available  = false;
    void*   m_pFactory   = nullptr; // ISpellCheckerFactory*
    void*   m_pChecker   = nullptr; // ISpellChecker*
    HMODULE m_hSpell     = nullptr;
};
