#include "SpellChecker.h"
#include <spellcheck.h>
#include <wrl/client.h>

using Microsoft::WRL::ComPtr;

SpellChecker::SpellChecker() {
    Init();
}

SpellChecker::~SpellChecker() {
    if (m_pChecker) {
        reinterpret_cast<ISpellChecker*>(m_pChecker)->Release();
        m_pChecker = nullptr;
    }
    if (m_pFactory) {
        reinterpret_cast<ISpellCheckerFactory*>(m_pFactory)->Release();
        m_pFactory = nullptr;
    }
}

void SpellChecker::Init() {
    // ISpellCheckerFactory is available on Windows 8+
    ISpellCheckerFactory* pFactory = nullptr;
    HRESULT hr = CoCreateInstance(__uuidof(SpellCheckerFactory), nullptr,
                                  CLSCTX_INPROC_SERVER,
                                  __uuidof(ISpellCheckerFactory),
                                  reinterpret_cast<void**>(&pFactory));
    if (FAILED(hr) || !pFactory) {
        m_available = false;
        return;
    }
    m_pFactory  = pFactory;
    m_available = true;

    // Default to system UI language
    LANGID lid = GetUserDefaultUILanguage();
    wchar_t langTag[LOCALE_NAME_MAX_LENGTH];
    LCIDToLocaleName(lid, langTag, LOCALE_NAME_MAX_LENGTH, 0);
    SetLanguage(langTag);
}

bool SpellChecker::SetLanguage(const std::wstring& langTag) {
    if (!m_available) return false;
    auto* pFactory = reinterpret_cast<ISpellCheckerFactory*>(m_pFactory);

    BOOL supported = FALSE;
    if (FAILED(pFactory->IsSupported(langTag.c_str(), &supported)) || !supported)
        return false;

    if (m_pChecker) {
        reinterpret_cast<ISpellChecker*>(m_pChecker)->Release();
        m_pChecker = nullptr;
    }

    ISpellChecker* pChecker = nullptr;
    HRESULT hr = pFactory->CreateSpellChecker(langTag.c_str(), &pChecker);
    if (FAILED(hr) || !pChecker) return false;

    m_pChecker = pChecker;
    return true;
}

bool SpellChecker::Check(const std::wstring& word) {
    if (!m_available || !m_pChecker || word.empty()) return true;
    auto* pChecker = reinterpret_cast<ISpellChecker*>(m_pChecker);

    IEnumSpellingError* pErrors = nullptr;
    if (FAILED(pChecker->Check(word.c_str(), &pErrors)) || !pErrors) return true;

    ISpellingError* pError = nullptr;
    bool hasError = SUCCEEDED(pErrors->Next(&pError)) && pError;
    if (pError) pError->Release();
    pErrors->Release();
    return !hasError;
}

std::vector<SpellSuggestion> SpellChecker::Suggest(const std::wstring& word) {
    std::vector<SpellSuggestion> results;
    if (!m_available || !m_pChecker || word.empty()) return results;
    auto* pChecker = reinterpret_cast<ISpellChecker*>(m_pChecker);

    IEnumString* pSuggestions = nullptr;
    if (FAILED(pChecker->Suggest(word.c_str(), &pSuggestions)) || !pSuggestions)
        return results;

    LPOLESTR suggestion = nullptr;
    while (pSuggestions->Next(1, &suggestion, nullptr) == S_OK) {
        results.push_back({ suggestion });
        CoTaskMemFree(suggestion);
    }
    pSuggestions->Release();
    return results;
}

std::vector<SpellChecker::Mistake> SpellChecker::Check(const std::wstring& text,
                                                        int startOffset) {
    std::vector<Mistake> mistakes;
    if (!m_available || !m_pChecker || text.empty()) return mistakes;
    auto* pChecker = reinterpret_cast<ISpellChecker*>(m_pChecker);

    IEnumSpellingError* pErrors = nullptr;
    if (FAILED(pChecker->Check(text.c_str(), &pErrors)) || !pErrors) return mistakes;

    ISpellingError* pError = nullptr;
    while (SUCCEEDED(pErrors->Next(&pError)) && pError) {
        ULONG start = 0, length = 0;
        pError->get_StartIndex(&start);
        pError->get_Length(&length);
        CORRECTIVE_ACTION action;
        pError->get_CorrectiveAction(&action);
        if (action != CORRECTIVE_ACTION_NONE) {
            mistakes.push_back({ startOffset + static_cast<int>(start),
                                 static_cast<int>(length) });
        }
        pError->Release();
        pError = nullptr;
    }
    pErrors->Release();
    return mistakes;
}

void SpellChecker::AddWord(const std::wstring& word) {
    if (!m_available || !m_pChecker || word.empty()) return;
    reinterpret_cast<ISpellChecker*>(m_pChecker)->Add(word.c_str());
}

void SpellChecker::IgnoreWord(const std::wstring& word) {
    if (!m_available || !m_pChecker || word.empty()) return;
    reinterpret_cast<ISpellChecker*>(m_pChecker)->Ignore(word.c_str());
}
