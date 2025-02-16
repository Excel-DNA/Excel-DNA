//  Copyright (c) Govert van Drimmelen. All rights reserved.

#include "loader.h"

#include "exports.h"
#include "host.h"
#include "utils.h"

HMODULE hModuleCurrent;
std::wstring xllPath;
// Flag to coordinate load/unload close and remove.
HMODULE lockModule;
bool locked = false;

// LoaderInitialize is called when the .dll gets PROCESS_ATTACH
// First initialization comes here.
// For now we only store our own module handle.
void LoaderInitialize(HMODULE hModule)
{
	hModuleCurrent = hModule;
}

// LoaderUnload is called when the .dll gets PROCESS_DETACH.
// Last chance to clean up anything.
// We just delete the temp .config file if we created one.
void LoaderUnload(bool processTerminating)
{
	// TODO:	tempConfig.destroy();
}

// Ensure that the library stays loaded.
// May be called many times, but should keep opened only until unlocked once.
void LockModule()
{
	if (!locked)
	{
		std::wstring xllFileName = xllPath;
		StripPath(xllFileName);
		lockModule = LoadLibrary(xllFileName.c_str());
		locked = true;
	}
}

// Allow the library to be unloaded.
void UnlockModule()
{
	if (locked)
	{
		FreeLibrary(lockModule);
		locked = false;
	}
}


// Ensure the runtime is loaded, get the managed loader running, and do the ExportInfo hook-up.
bool XlLibraryInitialize(XlAddInExportInfo* pExportInfo)
{
	std::wstring xllPath = GetAddInFullPath();
	WCHAR basePath[MAX_PATH] = { 0 };
	WCHAR drive[_MAX_DRIVE] = { 0 };
	_wsplitpath_s(xllPath.c_str(), drive, _MAX_DRIVE, basePath, MAX_PATH, NULL, 0, NULL, 0);
	int result = load_and_run(std::wstring(drive) + basePath, pExportInfo, hModuleCurrent, xllPath.c_str());

	return result == EXIT_SUCCESS;
	//
	//	HRESULT hr;
	//	ICorRuntimeHostPtr pHost;
	//	std::wstring clrVersion;
	//	bool shadowCopyFiles;
	//
	//	hr = GetClrOptions(clrVersion, shadowCopyFiles);
	//	if (FAILED(hr))
	//	{
	//		// SelectClrVersion shows diagnostic MessageBoxes if needed.
	//		// Perhaps remember that we are not loaded?
	//		return 0;
	//	}
	//
	//	bool allowedVersion = CompareNoCase(clrVersion, L"v4.0") >= 0;
	//	if (!allowedVersion)
	//	{
	//		ShowMessage(IDS_MSG_HEADER_NET4,
	//			IDS_MSG_BODY_NET4,
	//			IDS_MSG_FOOTER_ENSUREVERSION,
	//			hr);
	//		return 0;
	//	}
	//
	//	hr = LoadClr(clrVersion, &pHost);
	//	if (FAILED(hr) || pHost == NULL)
	//	{
	//		// LoadClr shows diagnostic MessageBoxes if needed.
	//		// Perhaps remember that we are not loaded?
	//		return 0;
	//	}
	//
	//	// If all is fine now, also start the CLR (always safe to do again).
	//	hr = pHost->Start();
	//	if (FAILED(hr))
	//	{
	//		ShowMessage(IDS_MSG_HEADER_NEEDCLR45,
	//			IDS_MSG_BODY_HOSTSTART,
	//			IDS_MSG_FOOTER_UNEXPECTED,
	//			hr);
	//		return 0;
	//	}
	//
	//	// Load (or find) the AppDomain that will contain the add-in
	//	std::wstring addInFullPath = GetAddInFullPath();
	//	_AppDomainPtr pAppDomain;
	//	_AssemblyPtr pLoaderAssembly;
	//	bool unloadAppDomain;
	//
	//	hr = LoadAppDomain(pHost, addInFullPath, createSandboxedAppDomain, shadowCopyFiles, pLoaderAssembly, pAppDomain, unloadAppDomain);
	//	if (FAILED(hr))
	//	{
	//		// Message already shown by LoadAppDomain
	//		return 0;
	//	}
	//
	//	_TypePtr pXlAddInType;
	//	hr = pLoaderAssembly->GetType_2(_bstr_t(L"ExcelDna.Loader.XlAddIn"), &pXlAddInType);
	//	if (FAILED(hr) || pXlAddInType == NULL)
	//	{
	//		ShowMessage(IDS_MSG_HEADER_APPDOMAIN,
	//			IDS_MSG_BODY_XLADDIN,
	//			IDS_MSG_FOOTER_UNEXPECTED,
	//			hr);
	//		return 0;
	//	}
	//
	//	SafeVariantArray initArgs(3);
	//	initArgs.lock();
	//#ifndef _M_X64
	//	initArgs.setElement(0, (INT32)pExportInfo);
	//	initArgs.setElement(1, (INT32)hModuleCurrent);
	//#else
	//	initArgs.setElement(0, (INT64)pExportInfo);
	//	initArgs.setElement(1, (INT64)hModuleCurrent);
	//#endif
	//	initArgs.setElement(2, SysAllocStringLen(addInFullPath.c_str(), static_cast<UINT>(addInFullPath.length())));
	//	initArgs.unlock();
	//	_variant_t initRetVal;
	//	_variant_t target;
	//#ifndef _M_X64
	//	hr = pXlAddInType->InvokeMember_3(_bstr_t("Initialize32"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
	//#else
	//	hr = pXlAddInType->InvokeMember_3(_bstr_t("Initialize64"), (BindingFlags)(BindingFlags_Static | BindingFlags_Public | BindingFlags_InvokeMethod), NULL, target, initArgs, &initRetVal);
	//#endif
	//	if (FAILED(hr) || initRetVal.boolVal == 0)
	//	{
	//		ShowMessage(IDS_MSG_HEADER_APPDOMAIN,
	//			IDS_MSG_BODY_XLADDININIT,
	//			IDS_MSG_FOOTER_CHECKINTEGRATION,
	//			hr);
	//		return 0;
	//	}
	//
	//	// Keep references needed for later host reference unload.
	//	if (unloadAppDomain)
	//	{
	//		pAppDomain_ForUnload = (IUnknown*)pAppDomain.Detach();
	//	}
	//	pHost_ForUnload = pHost.Detach();
	//
	//	return initRetVal.boolVal == 0 ? false : true;
}

// XlLibraryUnload is only called if we are unloading the add-in via the add-in manager,
// or when the add-in is re-loaded.
// TODO: Clean up the ALC ???
// TODO: Unhook the ExportInfo???
void XlLibraryUnload()
{
	//	if (pHost_ForUnload != NULL)
	//	{
	//		if (pAppDomain_ForUnload != NULL)
	//		{
	//			HRESULT hr = pHost_ForUnload->UnloadDomain(pAppDomain_ForUnload);
	//			pAppDomain_ForUnload->Release();
	//			pAppDomain_ForUnload = NULL;
	//			if (FAILED(hr))
	//			{
	//#if _DEBUG
	//				DebugBreak();
	//#endif
	//			}
	//		}
	//		//else
	//		//{
	//			// Unload according to the AppDomainId.
	//		//}
	//
	//		pHost_ForUnload->Release();
	//		pHost_ForUnload = NULL;
	//	}
		// Also delete the temp .config file, if we made one.
	// TODO: 	tempConfig.destroy();
}
