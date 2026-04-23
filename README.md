<div align="center">

<img src="resources/icon-preview.png" alt="WhatsUp Logo" width="120"/>

# What's Up
### 오픈소스 문서 편집기 · Open Source Text Editor

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078d7.svg)](https://github.com/elrang3843/whatsup)
[![Language](https://img.shields.io/badge/Language-C%2B%2B17-00599C.svg)](https://isocpp.org/)
[![Version](https://img.shields.io/badge/Version-1.0.0-brightgreen.svg)](https://github.com/elrang3843/whatsup/releases)

> 저사양 PC에서도 자유롭게 사용할 수 있는 오픈소스 한국어 문서 편집기입니다.  
> A lightweight, open-source document editor for everyone — even on low-spec PCs.

</div>

---

## 스크린샷 · Screenshots

| 메인 화면 | 시작 화면 |
|---|---|
| ![Main](resources/screenshot-main.png) | ![Splash](resources/screenshot-splash.png) |

---

## 특징 · Features

| 기능 | 설명 |
|---|---|
| **다국어 지원** | 한국어 / 영어 실시간 전환 |
| **다양한 형식** | HWP · HWPX · DOC · DOCX · MD · HTML · TXT |
| **맞춤법 검사** | Windows 8+ 내장 맞춤법 엔진 (ISpellChecker) |
| **자동 완성** | 문서 단어 학습 기반 팝업 완성 |
| **서식 도구** | 글자/문단 모양, 다단, 테두리/음영, 스타일 |
| **삽입 기능** | 표, 그림, 수식, 차트, 하이퍼링크, 책갈피, 이모지 등 |
| **인쇄** | Windows 기본 인쇄 다이얼로그 |
| **테마** | 라이트 / 다크 모드 |
| **저사양 최적화** | Win32 API 직접 사용 — 별도 런타임 불필요 |

---

## 지원 문서 형식 · Supported Formats

| 형식 | 읽기 | 쓰기 | 비고 |
|---|:---:|:---:|---|
| `.txt` | ✅ | ✅ | UTF-8 BOM, UTF-16 LE/BE, ANSI 자동 감지 |
| `.html` / `.htm` | ✅ | ✅ | 태그 제거 읽기, 완전한 HTML 출력 |
| `.md` | ✅ | ✅ | 마크다운 서식 제거 후 표시 |
| `.docx` | ✅ | ✅ | ZIP+XML 기반 (Word 2007+) |
| `.hwpx` | ✅ | ✅ | ZIP+XML 기반 (ISO/IEC 26300-3) |
| `.hwp` | ✅ | ❌ | HWP 5.x OLE2 바이너리 — 텍스트 추출만 지원 |
| `.doc` | ✅ | ❌ | Word 97-2003 OLE2 바이너리 — 텍스트 추출만 지원 |

---

## 메뉴 구성 · Menu Structure

```
파일(File)      새로 만들기 · 열기 · 저장 · 다른 이름으로 · 인쇄 · 닫기 · 종료
편집(Edit)      되돌리기 · 다시실행 · 오려두기 · 복사 · 붙이기 · 찾기 · 바꾸기 · 문서정보
입력(Insert)    글상자 · 문자표 · 도형 · 그림 · 이모지 · 표 · 차트 · 수식
                날짜/시간 · 주석 · 캡션 · 윗주 · 하이퍼링크 · 책갈피
서식(Format)    글자 모양 · 문단 모양 · 다단계 형식 · 개체 속성 · 테두리/음영 · 스타일
쪽(Page)        쪽 설정 · 머릿말 · 꼬리말 · 쪽번호 · 나누기 · 합치기
설정(Settings)  눈금/안내선 · 사용자 · 언어 · 테마
도움말(Help)    사용방법 · 라이선스 · WhatsUp 정보
```

---

## 빌드 방법 · Build Instructions

### 요구 사항 · Requirements

- **Windows** 7 이상 (권장: Windows 10/11)
- **CMake** 3.16 이상
- **컴파일러**: MSVC (Visual Studio 2019+) 또는 MinGW-w64
- **인터넷 연결**: 최초 빌드 시 zlib 자동 다운로드

### MSVC (Visual Studio)

```bat
git clone https://github.com/elrang3843/whatsup.git
cd whatsup
cmake -B build -G "Visual Studio 17 2022"
cmake --build build --config Release
```

### MinGW-w64

```bash
git clone https://github.com/elrang3843/whatsup.git
cd whatsup
cmake -B build -G "MinGW Makefiles"
cmake --build build
```

빌드 결과물: `build/WhatsUp.exe`

---

## 프로젝트 구조 · Project Structure

```
WhatsUp/
├── CMakeLists.txt               # 빌드 설정
├── LICENSE                      # Apache-2.0
├── resources/
│   └── manifest.xml             # DPI 인식, 비주얼 스타일
└── src/
    ├── main.cpp                 # WinMain 진입점
    ├── Application.h/cpp        # 앱 싱글톤, 설정 관리
    ├── MainWindow.h/cpp         # 메인 창 (메뉴, 툴바, 상태바)
    ├── Editor.h/cpp             # RichEdit 래퍼
    ├── Document.h/cpp           # 문서 모델
    ├── SplashScreen.h/cpp       # 시작 화면 (GDI 기어 로고)
    ├── resource.h               # 리소스 ID 정의
    ├── WhatsUp.rc               # 메뉴/다이얼로그 리소스
    ├── i18n/
    │   └── Localization.h/cpp   # 한국어 / 영어 문자열
    ├── formats/
    │   ├── IDocumentFormat.h    # 형식 핸들러 인터페이스
    │   ├── FormatManager.h/cpp  # 형식 자동 감지 및 라우팅
    │   ├── ZipReader.h/cpp      # 최소 ZIP 구현 (zlib 사용)
    │   ├── TxtFormat.h/cpp
    │   ├── HtmlFormat.h/cpp
    │   ├── MdFormat.h/cpp
    │   ├── DocxFormat.h/cpp
    │   ├── HwpxFormat.h/cpp
    │   ├── HwpFormat.h/cpp      # 읽기 전용
    │   └── DocFormat.h/cpp      # 읽기 전용
    ├── spell/
    │   └── SpellChecker.h/cpp   # ISpellCheckerFactory (Win8+)
    ├── autocomplete/
    │   └── AutoComplete.h/cpp   # 단어 완성 팝업
    └── dialogs/
        ├── AboutDialog.h/cpp
        ├── SettingsDialog.h/cpp
        ├── DocInfoDialog.h/cpp
        ├── InsertTableDialog.h/cpp
        ├── InsertHyperlinkDialog.h/cpp
        ├── InsertDateTimeDialog.h/cpp
        └── ...
```

---

## 라이선스 · License

```
Copyright 2026 Handtech

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0
```

이 프로젝트는 **Apache License 2.0** 하에 배포됩니다.  
시스템에 설치된 폰트를 사용하며, 별도의 저작권 침해 요소가 없습니다.

---

## 기여 · Contributing

버그 리포트, 기능 제안, Pull Request 모두 환영합니다.

1. 이 저장소를 Fork하세요.
2. 기능 브랜치를 만드세요: `git checkout -b feature/my-feature`
3. 변경 사항을 커밋하세요: `git commit -m "Add my feature"`
4. 브랜치에 Push하세요: `git push origin feature/my-feature`
5. Pull Request를 열어주세요.

---

<div align="center">

**Developed by [Handtech](https://github.com/elrang3843)**  
© 2026 Handtech. All rights reserved.  
Open Source · Apache-2.0

</div>
