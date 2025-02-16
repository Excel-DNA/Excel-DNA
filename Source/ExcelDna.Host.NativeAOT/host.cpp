//  Copyright (c) Govert van Drimmelen. All rights reserved.

#include "host.h"
#include "exports.h"
#include "TempDir.h"
#include "utils.h"
#include "dnainfo.h"

#include <filesystem>

TempDir tempDir(L"ExcelDna.Host.NativeAOT");

int load_and_run(const std::wstring& basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll)
{
	std::wstring hostFile(GetAddInFullPath());
	RenameExtension(hostFile, L".dll");

	if (!std::filesystem::exists(hostFile))
	{
		std::wstring hostFileName = hostFile;
		StripPath(hostFileName);

		hostFile = PathCombine(tempDir.GetPath(), hostFileName);
		if (!std::filesystem::exists(hostFile))
		{
			int r = WriteResourceToFile(hModuleXll, L"__MAIN__", L"NATIVE_ASSEMBLY", hostFile);
			if (r != EXIT_SUCCESS)
				return r;
		}
	}

	HINSTANCE handle = LoadLibrary(hostFile.c_str());

	if (handle == NULL)
	{
		ShowHostError(L"Loading " + hostFile + L" library failed.");
		return EXIT_FAILURE;
	}

	typedef short(__stdcall* xladdin_initialize_native_fn)(void* xlAddInExportInfo, void* hModuleXLL, void* pPathXLL, BYTE disableAssemblyContextUnload, void* pTempDirPath);

	xladdin_initialize_native_fn init = (xladdin_initialize_native_fn)GetProcAddress(handle, "Initialize");
	if (init == NULL)
	{
		ShowHostError(L"GetProcAddress Initialize failed.");
		return EXIT_FAILURE;
	}

	std::wstring tempDirPath = tempDir.GetPath();
	short res = init(pExportInfo, hModuleXll, (void*)pathXll, false, (void*)tempDirPath.c_str());

	return res == 0 ? EXIT_FAILURE : EXIT_SUCCESS;
}
