//  Copyright (c) Govert van Drimmelen. All rights reserved.

// xllhost.cpp : Defines the entry point for the XLL host
#include "host.h"
#include "exports.h"

// Include the official nethost API and indicate
// consumption should be as a static library.
#define NETHOST_USE_AS_STATIC
#include <nethost.h>

#include <coreclr_delegates.h>
#include <hostfxr.h>
#include <error_codes.h>

#include <cassert>
#include <iostream>
#include <string>
#include <filesystem>

#include "TempDir.h"
#include "utils.h"
#include "dnainfo.h"

using string_t = std::basic_string<char_t>;

#if _WIN64
const std::wstring requiredBitness = L"x64";
#else
const std::wstring requiredBitness = L"x86";
#endif

// Globals to hold hostfxr exports
hostfxr_initialize_for_runtime_config_fn init_fptr;
hostfxr_get_runtime_delegate_fn get_delegate_fptr;
hostfxr_get_runtime_property_value_fn get_property_fptr;
hostfxr_close_fn close_fptr;

// Forward declarations
bool load_hostfxr(int& rc, std::wstring& loadError);
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly(HMODULE hModuleXll, int majorRuntimeVersion, const std::wstring& rollForward, const std::wstring& runtimeFrameworkVersion);

// Provide a callback for any catastrophic failures.
// The provided callback will be the last call prior to a rude-abort of the process.
// void __stdcall set_failure_callback(failure_fn cb) {}

// Preload the runtime.
// The runtime is lazily loaded whenever the first export is called. This function
// preloads the runtime independent of calling any export and avoids the startup
// cost associated with calling an export for the first time.
void __stdcall preload_runtime(void) {}

TempDir tempDir(L"ExcelDna.Host");

// TODO: Might return the fn*
int load_and_run(const std::wstring& basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll)
{
	//
	// STEP 1: Load HostFxr and get exported hosting functions
	//
	int rc = 0;
	std::wstring loadError;
	if (!load_hostfxr(rc, loadError))
	{
		std::wstring msg;
		if (rc == CoreHostLibMissingFailure)
		{
			std::wstring requiredRuntime = L".NET Desktop Runtime 6.0.2+ " + requiredBitness;
			msg = std::format(L"{0} is not installed, corrupted or incomplete.\n\nYou can download {0} from https://dotnet.microsoft.com/en-us/download/dotnet/6.0", requiredRuntime);
		}
		else
		{
			msg = loadError;
		}
		ShowHostError(msg);

		return EXIT_FAILURE;
	}

	//
	// STEP 2: Initialize and start the .NET Core runtime
	//
	int majorRuntimeVersion;
	HRESULT hr = GetMajorRuntimeVersion(majorRuntimeVersion);
	if (FAILED(hr))
		majorRuntimeVersion = 6;

	std::wstring rollForward;
	hr = GetRollForward(rollForward);
	if (FAILED(hr))
		rollForward = L"";

	std::wstring runtimeFrameworkVersion;
	hr = GetRuntimeFrameworkVersion(runtimeFrameworkVersion);
	if (FAILED(hr))
		runtimeFrameworkVersion = L"";

	load_assembly_and_get_function_pointer_fn load_assembly_and_get_function_pointer = get_dotnet_load_assembly(hModuleXll, majorRuntimeVersion, rollForward, runtimeFrameworkVersion);
	if (load_assembly_and_get_function_pointer == nullptr)
		return EXIT_FAILURE;

	//
	// STEP 3: Copy managed assembly from resources to some temp file if needed
	//
	std::wstring hostFile = PathCombine(basePath, L"ExcelDna.ManagedHost.dll");
	if (!std::filesystem::exists(hostFile))
	{
		hostFile = PathCombine(tempDir.GetPath(), L"ExcelDna.ManagedHost.dll");
		if (!std::filesystem::exists(hostFile))
		{
			int r = WriteResourceToFile(hModuleXll, L"EXCELDNA.MANAGEDHOST", L"ASSEMBLY", hostFile);
			if (r != EXIT_SUCCESS)
				return r;
		}
	}

	//
	// STEP 4: Load managed assembly and get function pointer to a managed method
	//
	const string_t dotnetlib_path = hostFile;
	const char_t* dotnet_type = L"ExcelDna.ManagedHost.AddInInitialize, ExcelDna.ManagedHost";
	const char_t* dotnet_type_method = L"Initialize";

	// Function pointer to managed delegate with non-default signature
	typedef short (CORECLR_DELEGATE_CALLTYPE* xladdin_initialize_fn)(void* xlAddInExportInfo, void* hModuleXLL, void* pPathXLL, BYTE disableAssemblyContextUnload, void* pTempDirPath);
	xladdin_initialize_fn init = nullptr;
	rc = load_assembly_and_get_function_pointer(
		dotnetlib_path.c_str(),
		dotnet_type,
		dotnet_type_method,
		UNMANAGEDCALLERSONLY_METHOD,
		nullptr,
		(void**)&init);
	if (rc != 0 || init == nullptr)
	{
		std::wstringstream stream;
		stream << L"Loading ExcelDna.ManagedHost failed: " << std::hex << std::showbase << rc;
		ShowHostError(stream.str());
		return EXIT_FAILURE;
	}

	bool disableAssemblyContextUnload;
	hr = GetDisableAssemblyContextUnload(disableAssemblyContextUnload);
	if (FAILED(hr))
		disableAssemblyContextUnload = false;

	std::wstring tempDirPath = tempDir.GetPath();
	short res = init(pExportInfo, hModuleXll, (void*)pathXll, disableAssemblyContextUnload, (void*)tempDirPath.c_str());

	return res == 0 ? EXIT_FAILURE : EXIT_SUCCESS;
}


/********************************************************************************************
 * Function used to load and activate .NET Core
 ********************************************************************************************/

void* load_library(const char_t* path)
{
	HMODULE h = ::LoadLibraryW(path);
	return (void*)h;
}

void* get_export(void* h, const char* name)
{
	void* f = ::GetProcAddress((HMODULE)h, name);
	return f;
}

// Using the nethost library, discover the location of hostfxr and get exports
bool load_hostfxr(int& rc, std::wstring& loadError)
{
	// Pre-allocate a large buffer for the path to hostfxr
	char_t buffer[MAX_PATH];
	size_t buffer_size = sizeof(buffer) / sizeof(char_t);
	rc = get_hostfxr_path(buffer, &buffer_size, nullptr);
	if (rc != 0)
	{
		std::wstringstream stream;
		stream << "Getting hostfxr path failed.\n\nError code: " << std::hex << std::showbase << rc << ".";
		loadError = stream.str();
		return false;
	}

	// Load hostfxr and get desired exports
	void* lib = load_library(buffer);
	if (lib == nullptr)
	{
		loadError = std::format(L"Loading library {} failed.", buffer);
		return false;
	}

	init_fptr = (hostfxr_initialize_for_runtime_config_fn)get_export(lib, "hostfxr_initialize_for_runtime_config");
	if (init_fptr == nullptr)
	{
		loadError = L"Retrieving the address of the exported function hostfxr_initialize_for_runtime_config failed.";
		return false;
	}

	get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)get_export(lib, "hostfxr_get_runtime_delegate");
	if (get_delegate_fptr == nullptr)
	{
		loadError = L"Retrieving the address of the exported function hostfxr_get_runtime_delegate failed.";
		return false;
	}

	get_property_fptr = (hostfxr_get_runtime_property_value_fn)get_export(lib, "hostfxr_get_runtime_property_value");
	if (get_property_fptr == nullptr)
	{
		loadError = L"Retrieving the address of the exported function hostfxr_get_runtime_property_value failed.";
		return false;
	}

	close_fptr = (hostfxr_close_fn)get_export(lib, "hostfxr_close");
	if (close_fptr == nullptr)
	{
		loadError = L"Retrieving the address of the exported function hostfxr_close failed.";
		return false;
	}

	return true;
}

std::wstring get_runtime_property(const hostfxr_handle host_context_handle, const std::wstring& name)
{
	const wchar_t* value = NULL;
	int rc = get_property_fptr(host_context_handle, name.c_str(), &value);
	if (STATUS_CODE_SUCCEEDED(rc) && value != NULL)
	{
		return value;
	}

	return L"";
}

std::wstring get_loaded_runtime_version()
{
	return GetDirectoryName(get_runtime_property(NULL, L"FX_DEPS_FILE"));
}

// Load and initialize .NET Core and get desired function pointer for scenario
load_assembly_and_get_function_pointer_fn get_dotnet_load_assembly(HMODULE hModuleXll, int majorRuntimeVersion, const std::wstring& rollForward, const std::wstring& runtimeFrameworkVersion)
{
	std::wstring configFile = PathCombine(tempDir.GetPath(), L"ExcelDna.Host.runtimeconfig.json");
	std::string version(runtimeFrameworkVersion.length() > 0 ? ANSIWStringToString(runtimeFrameworkVersion) :
		(majorRuntimeVersion >= 7 ? std::format("{0}.0.0", majorRuntimeVersion) : "6.0.2"));

	std::wstring customRuntimeConfiguration;
	bool useCustomRuntimeConfiguration = (SUCCEEDED(GetCustomRuntimeConfiguration(customRuntimeConfiguration)) && !customRuntimeConfiguration.empty());
	if (!useCustomRuntimeConfiguration)
	{
		std::string tfm(std::format("net{0}.0", majorRuntimeVersion));
		std::string rollForwardOption = rollForward.length() > 0 ? std::format(R"("rollForward": "{0}",)", ANSIWStringToString(rollForward)) : "";

		std::string configText = std::format(R"({{
  "runtimeOptions": {{
    "tfm": "{0}",
    {2}
    "framework": {{
      "name": "Microsoft.WindowsDesktop.App",
      "version": "{1}"
    }}
  }}
}})", tfm, version, rollForwardOption);
		HRESULT hr = WriteAllBytes(configFile, (void*)configText.c_str(), static_cast<DWORD>(configText.length()));
		if (FAILED(hr))
		{
			std::wstringstream stream;
			stream << "Saving ExcelDna.Host.runtimeconfig.json failed: " << std::hex << std::showbase << hr;
			ShowHostError(stream.str());
			return nullptr;
		}
	}
	else
	{
		std::wstring customRuntimeConfigurationFilePath = PathCombine(GetDirectory(GetAddInFullPath()), customRuntimeConfiguration);
		if (std::filesystem::exists(customRuntimeConfigurationFilePath))
		{
			configFile = customRuntimeConfigurationFilePath;
		}
		else
		{
			int r = WriteResourceToFile(hModuleXll, L"__CUSTOM_RUNTIMECONFIG__", L"SOURCE", configFile);
			if (r != EXIT_SUCCESS)
				return nullptr;
		}
	}

	// Load .NET Core
	void* load_assembly_and_get_function_pointer = nullptr;
	hostfxr_handle cxt = nullptr;
	int rc = init_fptr(configFile.c_str(), nullptr, &cxt);
	if (!STATUS_CODE_SUCCEEDED(rc) || cxt == nullptr)
	{
		std::wstring requiredRuntime = (useCustomRuntimeConfiguration ? L"custom .NET runtime" : std::format(L".NET Desktop Runtime {0} {1}", std::wstring(version.begin(), version.end()), requiredBitness));
		if (rc == CoreHostIncompatibleConfig)
		{
			std::wstring msg = L"The required " + requiredRuntime + L" is incompatible with the runtime " + get_loaded_runtime_version() + L" already loaded in the process.\n\nYou can try to disable other Excel add-ins to resolve the conflict.";
			ShowHostError(msg);
		}
		else if (rc == FrameworkMissingFailure)
		{
			std::wstring msg = std::format(L"It was not possible to find a compatible framework version for {0}.\n\nYou can download {0} from https://dotnet.microsoft.com/en-us/download/dotnet/", requiredRuntime);
			ShowHostError(msg);
		}
		else
		{
			std::wstringstream stream;
			stream << L".NET runtime initialization failed: " << std::hex << std::showbase << rc << std::endl << std::endl;
			stream << L"You can find more information about this error at https://github.com/dotnet/runtime/blob/main/docs/design/features/host-error-codes.md";
			ShowHostError(stream.str());
		}
		close_fptr(cxt);
		return nullptr;
	}

	// Get the load assembly function pointer
	rc = get_delegate_fptr(
		cxt,
		hdt_load_assembly_and_get_function_pointer,
		&load_assembly_and_get_function_pointer);
	if (!STATUS_CODE_SUCCEEDED(rc) || load_assembly_and_get_function_pointer == nullptr)
	{
		std::wstringstream stream;
		stream << "Get .NET runtime delegate failed: " << std::hex << std::showbase << rc << std::endl << std::endl;
		stream << L"You can find more information about this error at https://github.com/dotnet/runtime/blob/main/docs/design/features/host-error-codes.md";
		ShowHostError(stream.str());
		close_fptr(cxt);
		return nullptr;
	}

	close_fptr(cxt);
	return (load_assembly_and_get_function_pointer_fn)load_assembly_and_get_function_pointer;
}
