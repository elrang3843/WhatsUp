#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <objbase.h>

#include "Application.h"
#include "SplashScreen.h"

int WINAPI WinMain(
    _In_     HINSTANCE hInstance,
    _In_opt_ HINSTANCE /*hPrevInstance*/,
    _In_     LPSTR     /*lpCmdLine*/,
    _In_     int       nCmdShow)
{
    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr))
        return -1;

    SplashScreen::Show(hInstance, 2500);

    int result = -1;
    if (Application::Instance().Init(hInstance, nCmdShow))
        result = Application::Instance().Run();

    CoUninitialize();
    return result;
}
