#pragma once

#include <Windows.h>

// EXPORT_COUNT defines the size of the thunk table. 
// Must match the number of functions exported in exports32.cpp and exports64.asm
constexpr INT32 EXPORT_COUNT = 10000;

// The function pointers will be exported from managed code in the StdCall convention.
typedef void(__stdcall* PFN)();
typedef short(__stdcall* PFN_SHORT_VOID)();
typedef void(__stdcall* PFN_VOID_LPXLOPER12)(void*);
typedef void* (__stdcall* PFN_LPXLOPER12_LPXLOPER12)(void*);
typedef void(__stdcall* PFN_PFNEXCEL12)(void*);
typedef void(__stdcall* PFN_VOID_DOUBLE)(double);
typedef void(__stdcall* PFN_VOID_VOID)();
typedef long(__stdcall* PFN_LPENHELPER)(int, void*);
typedef HRESULT(__stdcall* PFN_GET_CLASS_OBJECT)(CLSID clsid, IID iid, LPVOID* ppv);
typedef HRESULT(__stdcall* PFN_HRESULT_VOID)();

// ExcelDna add-ins do not implement xlAutoRegister because all
// registrations contain the signature from the start.
struct XlAddInExportInfo
{
	INT32 ExportInfoVersion;
	DWORD AppDomainId;
	PFN_SHORT_VOID				pXlAutoOpen;
	PFN_SHORT_VOID				pXlAutoClose;
	PFN_SHORT_VOID				pXlAutoRemove;
	PFN_VOID_LPXLOPER12			pXlAutoFree12;
	PFN_PFNEXCEL12				pSetExcel12EntryPt;
	PFN_HRESULT_VOID			pDllRegisterServer;
	PFN_HRESULT_VOID			pDllUnregisterServer;
	PFN_GET_CLASS_OBJECT		pDllGetClassObject;
	PFN_HRESULT_VOID			pDllCanUnloadNow;
	PFN_VOID_DOUBLE				pSyncMacro;
	PFN_LPXLOPER12_LPXLOPER12	pRegistrationInfo;
	PFN_VOID_VOID				pCalculationCanceled;
	PFN_VOID_VOID				pCalculationEnded;
	PFN_LPENHELPER				pLPenHelper;
	// The thunk table that hooks up the fxxx exports from the .xll with the marshaled function pointers.
	INT32  ThunkTableLength;
	PFN* ThunkTable;           // Actually (PFN ThunkTable[EXPORT_COUNT])
};

XlAddInExportInfo* CreateExportInfo();
