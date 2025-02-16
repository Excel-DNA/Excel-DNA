// dllmain.cpp : Defines the entry point for the DLL application.
#include "loader.h"

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        LoaderInitialize(hModule);
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        LoaderUnload(lpReserved != nullptr);
        break;
    }
    return TRUE;
}

