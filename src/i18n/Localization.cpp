#include "Localization.h"
#include <cassert>

Language Localization::s_lang = Language::Korean;

// Korean strings — keep in SID order
const wchar_t* const Localization::s_ko[] = {
    // APP
    L"WhatsUp",
    L"1.0.0",
    L"© 2026 Handtech. All rights reserved.",
    L"제목 없음",

    // MENU_FILE
    L"파일(&F)",
    L"새로 만들기(&N)\tAlt+N",
    L"열기(&O)...\tAlt+O",
    L"저장(&S)\tAlt+S",
    L"다른 이름으로 저장(&V)...\tAlt+V",
    L"인쇄(&P)...\tAlt+P",
    L"닫기\tCtrl+F4",
    L"종료(&X)\tAlt+X",

    // MENU_EDIT
    L"편집(&E)",
    L"되돌리기(&U)\tCtrl+Z",
    L"다시 실행(&R)\tCtrl+Shift+Z",
    L"오려두기(&T)\tCtrl+X",
    L"복사하기(&C)\tCtrl+C",
    L"붙이기(&P)\tCtrl+V",
    L"찾기(&F)...\tCtrl+F",
    L"바꾸기(&H)...\tCtrl+H",
    L"문서 정보(&I)\tCtrl+Q",
    L"모두 선택(&A)\tCtrl+A",

    // MENU_INPUT
    L"입력(&I)",
    L"글상자(&B)",
    L"문자표(&M)...",
    L"도형(&S)",
    L"그림(&G)...",
    L"이모지(&E)...",
    L"표(&T)...",
    L"차트(&C)...",
    L"수식(&F)...",
    L"날짜/시간(&D)...",
    L"주석(&N)...",
    L"캡션(&A)...",
    L"윗주(&H)...",
    L"하이퍼링크(&L)...",
    L"책갈피(&K)...",

    // MENU_FORMAT
    L"서식(&O)",
    L"글자 모양(&C)...",
    L"문단 모양(&P)...",
    L"다단계 형식(&M)...",
    L"개체 속성(&O)...",
    L"테두리/음영(&B)...",
    L"스타일(&S)...",

    // MENU_PAGE
    L"쪽(&G)",
    L"쪽 설정(&S)...",
    L"머릿말(&H)...",
    L"꼬리말(&F)...",
    L"쪽 번호(&N)...",
    L"나누기(&D)",
    L"합치기(&M)",

    // MENU_SETTINGS
    L"설정(&T)",
    L"눈금/안내선(&G)...",
    L"사용자(&U)...",
    L"언어(&L)...",
    L"테마(&H)...",

    // MENU_HELP
    L"도움말(&H)",
    L"사용 방법(&M)...",
    L"라이선스(&L)...",
    L"WhatsUp 정보(&A)...",

    // DLG
    L"확인",
    L"취소",
    L"예",
    L"아니오",
    L"닫기",

    // MSG
    L"저장하지 않은 변경 사항이 있습니다. 저장하시겠습니까?",
    L"저장되지 않은 문서",
    L"파일을 찾을 수 없습니다.",
    L"파일을 열지 못했습니다.",
    L"파일을 저장하지 못했습니다.",
    L"인쇄에 실패했습니다.",
    L"지원하지 않는 파일 형식입니다.",
    L"HWP(.hwp) 이진 형식은 라이선스 문제로 쓰기를 지원하지 않습니다.\n다른 형식(HWPX, DOCX, HTML, TXT 등)으로 저장해 주세요.",
    L"DOC(.doc) 이진 형식은 라이선스 문제로 쓰기를 지원하지 않습니다.\n다른 형식(DOCX, HTML, TXT 등)으로 저장해 주세요.",
    L"맞춤법 검사 서비스를 사용할 수 없습니다 (Windows 8 이상 필요).",
    L"일반 텍스트(.txt)로 저장하면 글꼴 지정, 도형, 표, 서식 등 모든 서식 정보가 손실됩니다.\n계속 저장하시겠습니까?",
    L"서식 정보 손실 경고",

    // STATUS
    L"준비",
    L"수정됨",
    L"단어: %d",
    L"문자: %d",
    L"쪽: %d / %d",
    L"줄: %d",
    L"열: %d",

    // FILTER
    L"지원 문서 (*.hwp;*.hwpx;*.doc;*.docx;*.md;*.html;*.htm;*.txt)",
    L"일반 텍스트 (*.txt)",
    L"HTML 문서 (*.html;*.htm)",
    L"마크다운 (*.md)",
    L"Word 문서 (*.docx)",
    L"한글 XML 문서 (*.hwpx)",
    L"한글 문서 (*.hwp)",
    L"Word 97-2003 (*.doc)",
    L"모든 파일 (*.*)",

    // TABS
    L"일반",
    L"사용자",
    L"맞춤법",
    L"눈금",
    L"속성",
    L"통계",
    L"워터마크",

    // TOOLTIPS
    L"새로 만들기 (Alt+N)",
    L"열기 (Alt+O)",
    L"저장 (Alt+S)",
    L"인쇄 (Alt+P)",
    L"되돌리기 (Ctrl+Z)",
    L"다시 실행 (Ctrl+Shift+Z)",
    L"오려두기 (Ctrl+X)",
    L"복사하기 (Ctrl+C)",
    L"붙이기 (Ctrl+V)",
    L"굵게 (Ctrl+B)",
    L"기울임 (Ctrl+I)",
    L"밑줄 (Ctrl+U)",
    L"취소선",
    L"왼쪽 정렬",
    L"가운데 정렬",
    L"오른쪽 정렬",
    L"양쪽 정렬",
    L"들여쓰기",
    L"내어쓰기",
    L"글자 색",
    L"형광펜",
    L"찾기 (Ctrl+F)",
    L"바꾸기 (Ctrl+H)",
};

// English strings — keep in SID order
const wchar_t* const Localization::s_en[] = {
    // APP
    L"WhatsUp",
    L"1.0.0",
    L"© 2026 Handtech. All rights reserved.",
    L"Untitled",

    // MENU_FILE
    L"&File",
    L"&New\tAlt+N",
    L"&Open...\tAlt+O",
    L"&Save\tAlt+S",
    L"Save &As...\tAlt+V",
    L"&Print...\tAlt+P",
    L"&Close\tCtrl+F4",
    L"E&xit\tAlt+X",

    // MENU_EDIT
    L"&Edit",
    L"&Undo\tCtrl+Z",
    L"&Redo\tCtrl+Shift+Z",
    L"Cu&t\tCtrl+X",
    L"&Copy\tCtrl+C",
    L"&Paste\tCtrl+V",
    L"&Find...\tCtrl+F",
    L"&Replace...\tCtrl+H",
    L"Document &Info\tCtrl+Q",
    L"Select &All\tCtrl+A",

    // MENU_INPUT
    L"&Insert",
    L"Text &Box",
    L"&Character Map...",
    L"&Shapes",
    L"&Picture...",
    L"&Emoji...",
    L"&Table...",
    L"&Chart...",
    L"&Formula...",
    L"&Date/Time...",
    L"Co&mment...",
    L"C&aption...",
    L"&Headnote...",
    L"Hyper&link...",
    L"Boo&kmark...",

    // MENU_FORMAT
    L"F&ormat",
    L"&Character...",
    L"&Paragraph...",
    L"&Columns...",
    L"&Object Properties...",
    L"&Border/Shading...",
    L"&Style...",

    // MENU_PAGE
    L"&Page",
    L"Page &Setup...",
    L"&Header...",
    L"&Footer...",
    L"Page &Number...",
    L"&Split",
    L"&Merge",

    // MENU_SETTINGS
    L"&Settings",
    L"&Grid/Guidelines...",
    L"&User...",
    L"&Language...",
    L"&Theme...",

    // MENU_HELP
    L"&Help",
    L"&User Manual...",
    L"&Licenses...",
    L"&About WhatsUp...",

    // DLG
    L"OK",
    L"Cancel",
    L"Yes",
    L"No",
    L"Close",

    // MSG
    L"There are unsaved changes. Do you want to save?",
    L"Unsaved Document",
    L"File not found.",
    L"Failed to open file.",
    L"Failed to save file.",
    L"Print failed.",
    L"Unsupported file format.",
    L"Writing the HWP (.hwp) binary format is not supported due to licensing restrictions.\nPlease save in another format (HWPX, DOCX, HTML, TXT, etc.).",
    L"Writing the DOC (.doc) binary format is not supported due to licensing restrictions.\nPlease save in another format (DOCX, HTML, TXT, etc.).",
    L"Spell check service unavailable (requires Windows 8 or later).",
    L"Saving as plain text (.txt) will discard all formatting including fonts, shapes, tables, and styles.\nContinue saving?",
    L"Formatting Loss Warning",

    // STATUS
    L"Ready",
    L"Modified",
    L"Words: %d",
    L"Chars: %d",
    L"Page: %d / %d",
    L"Line: %d",
    L"Col: %d",

    // FILTER
    L"Supported Documents (*.hwp;*.hwpx;*.doc;*.docx;*.md;*.html;*.htm;*.txt)",
    L"Plain Text (*.txt)",
    L"HTML Document (*.html;*.htm)",
    L"Markdown (*.md)",
    L"Word Document (*.docx)",
    L"HWP XML Document (*.hwpx)",
    L"HWP Document (*.hwp)",
    L"Word 97-2003 (*.doc)",
    L"All Files (*.*)",

    // TABS
    L"General",
    L"User",
    L"Spell Check",
    L"Grid",
    L"Properties",
    L"Statistics",
    L"Watermark",

    // TOOLTIPS
    L"New (Alt+N)",
    L"Open (Alt+O)",
    L"Save (Alt+S)",
    L"Print (Alt+P)",
    L"Undo (Ctrl+Z)",
    L"Redo (Ctrl+Shift+Z)",
    L"Cut (Ctrl+X)",
    L"Copy (Ctrl+C)",
    L"Paste (Ctrl+V)",
    L"Bold (Ctrl+B)",
    L"Italic (Ctrl+I)",
    L"Underline (Ctrl+U)",
    L"Strikethrough",
    L"Align Left",
    L"Align Center",
    L"Align Right",
    L"Justify",
    L"Increase Indent",
    L"Decrease Indent",
    L"Font Color",
    L"Highlight",
    L"Find (Ctrl+F)",
    L"Replace (Ctrl+H)",
};

static_assert(static_cast<int>(StrID::_COUNT) == ARRAYSIZE(Localization::s_ko),
              "Korean string table size mismatch");
static_assert(static_cast<int>(StrID::_COUNT) == ARRAYSIZE(Localization::s_en),
              "English string table size mismatch");

void Localization::SetLanguage(Language lang) {
    s_lang = lang;
}

Language Localization::GetLanguage() {
    return s_lang;
}

const wchar_t* Localization::Get(StrID id) {
    int idx = static_cast<int>(id);
    if (idx < 0 || idx >= static_cast<int>(StrID::_COUNT)) return L"";
    return (s_lang == Language::Korean) ? s_ko[idx] : s_en[idx];
}

std::wstring Localization::GetFileFilter() {
    // File filter is pairs of description\0pattern\0, terminated by extra \0
    std::wstring filter;
    auto add = [&](StrID descId, const wchar_t* pattern) {
        filter += Get(descId);
        filter += L'\0';
        filter += pattern;
        filter += L'\0';
    };
    add(StrID::FILTER_ALL_DOCS,  L"*.hwp;*.hwpx;*.doc;*.docx;*.md;*.html;*.htm;*.txt");
    add(StrID::FILTER_HWPX,      L"*.hwpx");
    add(StrID::FILTER_HWP,       L"*.hwp");
    add(StrID::FILTER_DOCX,      L"*.docx");
    add(StrID::FILTER_DOC,       L"*.doc");
    add(StrID::FILTER_MD,        L"*.md");
    add(StrID::FILTER_HTML,      L"*.html;*.htm");
    add(StrID::FILTER_TXT,       L"*.txt");
    add(StrID::FILTER_ALL_FILES, L"*.*");
    filter += L'\0'; // double-null terminator
    return filter;
}
