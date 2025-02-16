//  Copyright (c) Govert van Drimmelen. All rights reserved.

// xllhost.cpp : Defines the entry point for the XLL host
#include "loader.h"
#include "exports.h"
#include "LPenHelper.h"
#include "dnainfo.h"

// Minimal parts of XLOPER types, 
// used only for xlAddInManagerInfo(12). Really.

struct XLOPER12
{
	union
	{
		double num;					/* xltypeNum */
		void* str;					/* xltypeStr */
		int err;					/* xltypeErr */
		struct
		{
			double unused1;
			double unused2;
			double unused3;
		} unused;
	} val;
	DWORD xltype; // Should be at offset 24 bytes
};
constexpr DWORD xltypeNum = 1;
constexpr DWORD xltypeStr = 2;
constexpr DWORD xltypeErr = 16;
constexpr DWORD xlerrValue = 15;

// The one and only ExportInfo
XlAddInExportInfo* pExportInfo = NULL;
bool removed = false;    // Used to check whether AutoRemove is called before AutoClose.
bool autoOpened = false; // Not set when loaded for COM server only. Used for re-open check.

// The actual thunk table 
extern "C"
{
	PFN thunks[EXPORT_COUNT];
}

XlAddInExportInfo* CreateExportInfo()
{
	pExportInfo = new XlAddInExportInfo();
	pExportInfo->ExportInfoVersion = 10;
	pExportInfo->pXlAutoOpen = NULL;
	pExportInfo->pXlAutoClose = NULL;
	pExportInfo->pXlAutoRemove = NULL;
	pExportInfo->pXlAutoFree12 = NULL;
	pExportInfo->pSetExcel12EntryPt = NULL;
	pExportInfo->pDllRegisterServer = NULL;
	pExportInfo->pDllUnregisterServer = NULL;
	pExportInfo->pDllGetClassObject = NULL;
	pExportInfo->pDllCanUnloadNow = NULL;
	pExportInfo->pSyncMacro = NULL;
	pExportInfo->pRegistrationInfo = NULL;
	pExportInfo->pCalculationCanceled = NULL;
	pExportInfo->pCalculationEnded = NULL;
	pExportInfo->pLPenHelper = LPenHelper;
	pExportInfo->ThunkTableLength = EXPORT_COUNT;
	pExportInfo->ThunkTable = (PFN*)thunks;
	return pExportInfo;
}

// Safe to be called repeatedly, but not from multiple threads
bool EnsureInitialized()
{
	if (pExportInfo != NULL)
	{
		return true;
	}
	XlAddInExportInfo* pExportInfoTemp = CreateExportInfo();
	bool initOK = XlLibraryInitialize(pExportInfoTemp);
	if (initOK)
	{
		pExportInfo = pExportInfoTemp;
		return true;
	}
	return false;
}

// Called only when AutoClose is called after AutoRemove.
void Uninitialize()
{
	delete pExportInfo;
	pExportInfo = NULL;
	for (int i = 0; i < EXPORT_COUNT; i++)
	{
		thunks[i] = NULL;
	}
}

// Forward declares, since these are now called by AutoOpen.
short __stdcall xlAutoClose();
short __stdcall xlAutoRemove();

// Excel Add-In standard exports
short __stdcall xlAutoOpen()
{
	short result = 0;

	// If we are loaded as an add-in already, then ensure re-load = AddInRemove + AutoClose + AutoOpen,
	// which mains a clean AppDomain for each load.
	if (autoOpened)
	{
		xlAutoRemove();
		xlAutoClose();
	}

	if (EnsureInitialized() &&
		pExportInfo->pXlAutoOpen != NULL)
	{
		result = pExportInfo->pXlAutoOpen();
		LockModule();
		// Set the 'removed' flag to false, which prevents AutoClose from actually unloading (or calling through to the add-in),
		// unless AutoRemove is called first (from the add-in manager, a host or the re-open sequence above).
		removed = false;
		// Keep track that we are loaded as an add-in, not just a COM or RTD server.
		// This allows us to re-open in a clean AppDomain, yet load COM server first then add-in without damage.
		autoOpened = true;
	}
	return result;
}

short __stdcall xlAutoClose()
{
	short result = 0;
	if (EnsureInitialized() &&
		pExportInfo->pXlAutoClose != NULL)
	{
		result = pExportInfo->pXlAutoClose();
		if (removed)
		{
			// TODO: Consider how and when to unload
			//       Unloading the AppDomain could be a bit too dramatic if we are serving as a COM Server or RTD Server directly.
			// DOCUMENT: What the current implementation is.
			// No more managed functions should be called.
			Uninitialize();

			// Complete the clean-up by unloading AppDomain
			XlLibraryUnload();
			// ... recording that we are no longer open as an add-in.
			autoOpened = false;
			// ...and allowing the .xll itself to be unloaded
			UnlockModule();
		}
	}
	return result;
}

// Since v0.29 loading is much more expensive, so I want to reduce the number of times we load.
// We've never used or exposed xlAutoAdd to Excel-DNA addins, so no harm in disabling for now.
// To add back, also uncomment in the ExcelDna.def file.
//short __stdcall xlAutoAdd()
//{
//	short result = 0;
//	if (EnsureInitialized() && 
//		pExportInfo->pXlAutoAdd != NULL)
//	{
//		result = pExportInfo->pXlAutoAdd();
//	}
//	return result;
//}

short __stdcall xlAutoRemove()
{
	short result = 0;
	if (EnsureInitialized() &&
		pExportInfo->pXlAutoRemove != NULL)
	{
		result = pExportInfo->pXlAutoRemove();
		// Set the 'removed' flag which will allow the AutoClose to actually unload (and call through to the add-in).
		removed = true;
	}
	return result;
}

void __stdcall xlAutoFree12(void* pXloper12)
{
	if (pExportInfo != NULL && pExportInfo->pXlAutoFree12 != NULL)
	{
		pExportInfo->pXlAutoFree12(pXloper12);
	}
}

XLOPER12* __stdcall xlAddInManagerInfo12(XLOPER12* pXloper)
{
	static XLOPER12 result;
	static wchar_t name[256];

	// Return error by default
	result.xltype = xltypeErr;
	result.val.err = xlerrValue;

	if (pXloper->xltype == 1 && pXloper->val.num == 1.0)
	{
		std::wstring addInName;
		HRESULT hr = GetAddInName(addInName);
		if (!FAILED(hr))
		{
			// We could probably use CString as is (maybe with truncation)!?
			int length = (int)min(addInName.length(), 254);
			name[0] = (wchar_t)length;
			wchar_t* pName = (wchar_t*)name + 1;
			lstrcpyn(pName, addInName.c_str(), length + 1);
			result.xltype = xltypeStr;
			result.val.str = name;
		}
	}

	return &result;
}

// Support for Excel 2010 SDK - used when loading under HPC XLL Host
void __stdcall SetExcel12EntryPt(void* pexcel12New)
{
	if (EnsureInitialized() &&
		pExportInfo->pSetExcel12EntryPt != NULL)
	{
		pExportInfo->pSetExcel12EntryPt(pexcel12New);
	}
}

// We are also a COM Server, to support the =RTD(...) worksheet function and VBA ComServer integration.
STDAPI DllRegisterServer()
{
	HRESULT result = E_UNEXPECTED;
	if (EnsureInitialized() &&
		pExportInfo->pDllRegisterServer != NULL)
	{
		result = pExportInfo->pDllRegisterServer();
	}
	return result;
}

STDAPI DllUnregisterServer()
{
	HRESULT result = E_UNEXPECTED;
	if (EnsureInitialized() &&
		pExportInfo->pDllUnregisterServer != NULL)
	{
		result = pExportInfo->pDllUnregisterServer();
	}
	return result;
}

HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID iid, void** ppv)
{
	HRESULT result = E_UNEXPECTED;
	GUID cls = clsid;
	GUID i = iid;
	if (EnsureInitialized() &&
		pExportInfo->pDllGetClassObject != NULL)
	{

		result = pExportInfo->pDllGetClassObject(cls, i, ppv);
	}
	return result;
}

HRESULT __stdcall DllCanUnloadNow()
{
	// CONSIDER: This caused problems for unloaded add-ins, when shutting Excel down.
	//           We need to add a flag that tracks whether the add-in has beren unloaded.
	//           Always returning FALSE is what was happening internally anyway, so we're no worse off than before.

	//HRESULT result = S_OK;
	//if (EnsureInitialized() &&
	//	pExportInfo->pDllCanUnloadNow != NULL)
	//{
	//	result = pExportInfo->pDllCanUnloadNow();
	//}
	//return result;

	return S_FALSE;
}

void __stdcall SyncMacro(double param)
{
	if (EnsureInitialized() &&
		pExportInfo->pSyncMacro != NULL)
	{
		pExportInfo->pSyncMacro(param);
	}
}

XLOPER12* __stdcall RegistrationInfo(XLOPER12* param)
{
	if (EnsureInitialized() &&
		pExportInfo->pRegistrationInfo != NULL)
	{
		return (XLOPER12*)pExportInfo->pRegistrationInfo(param);
	}
	return NULL;
}

void __stdcall CalculationCanceled()
{
	if (EnsureInitialized() &&
		pExportInfo->pCalculationCanceled != NULL)
	{
		pExportInfo->pCalculationCanceled();
	}
}

void __stdcall CalculationEnded()
{
	if (EnsureInitialized() &&
		pExportInfo->pCalculationEnded != NULL)
	{
		pExportInfo->pCalculationEnded();
	}
}
