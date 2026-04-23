cmake_minimum_required(VERSION 3.16)
project(WhatsUp VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Windows 7+ target
add_compile_definitions(
    UNICODE _UNICODE
    WINVER=0x0601 _WIN32_WINNT=0x0601
    _WIN32_IE=0x0700
    NOMINMAX
    WIN32_LEAN_AND_MEAN
)

# Fetch zlib for DOCX/HWPX ZIP support
include(FetchContent)
FetchContent_Declare(
    zlib
    GIT_REPOSITORY https://github.com/madler/zlib.git
    GIT_TAG        v1.3.1
    OVERRIDE_FIND_PACKAGE
)
FetchContent_MakeAvailable(zlib)

file(GLOB_RECURSE APP_SOURCES CONFIGURE_DEPENDS "src/*.cpp")
file(GLOB_RECURSE APP_HEADERS CONFIGURE_DEPENDS "src/*.h")

add_executable(WhatsUp WIN32
    ${APP_SOURCES}
    ${APP_HEADERS}
    src/WhatsUp.rc
)

target_include_directories(WhatsUp PRIVATE
    src/
    ${zlib_SOURCE_DIR}
    ${zlib_BINARY_DIR}
)

target_link_libraries(WhatsUp PRIVATE
    zlibstatic
    comctl32
    comdlg32
    gdi32
    gdiplus
    user32
    shell32
    ole32
    oleaut32
    uuid
    uxtheme
    imm32
    shlwapi
    winmm
    winspool
)

if(MSVC)
    target_compile_options(WhatsUp PRIVATE /W4 /utf-8 /wd4100 /wd4505 /wd4244)
    target_link_options(WhatsUp PRIVATE /SUBSYSTEM:WINDOWS /MANIFEST:NO)
    set_property(TARGET WhatsUp PROPERTY MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
    # RC compiler needs resources/ (for manifest.xml) and images/ (for splash PNG)
    set_source_files_properties(src/WhatsUp.rc PROPERTIES
        COMPILE_FLAGS "/I \"${CMAKE_CURRENT_SOURCE_DIR}/resources\" /I \"${CMAKE_CURRENT_SOURCE_DIR}/images\"")
endif()

if(MINGW)
    target_compile_options(WhatsUp PRIVATE -Wall -Wextra -Wno-unused-parameter)
    target_link_options(WhatsUp PRIVATE -mwindows -static-libgcc -static-libstdc++)
endif()

set_target_properties(WhatsUp PROPERTIES
    OUTPUT_NAME "WhatsUp"
    WIN32_EXECUTABLE TRUE
)
