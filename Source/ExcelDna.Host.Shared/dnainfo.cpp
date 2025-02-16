//  Copyright (c) Govert van Drimmelen. All rights reserved.

#include "dnainfo.h"
#include "utils.h"
#include "resource.h"

extern HMODULE hModuleCurrent;

HRESULT GetDnaHeader(bool showErrors, std::wstring& header)
{
	// We find the .dna file and load a 1k string from the file.
	// To locate the file:
	// 1. First check for packed __MAIN__ DNA resource,
	// 2. Else load file next to .xll file, 
	// Else E_FAIL.
	// This sequence matches the load sequence in ExcelDna.Integration.DnaLibrary.Initialize().
	// NOTE: __MAIN__ DNA resource can not currently be compressed.

	HRESULT hr = E_FAIL;
	const DWORD MAX_HEADER_LENGTH = 1024;
	DWORD headerLength = 0;
	BYTE headerBuffer[MAX_HEADER_LENGTH];

	HRSRC hResDna = FindResource(hModuleCurrent, L"__MAIN__", L"DNA");
	if (hResDna != NULL)
	{
		HGLOBAL hDna = LoadResource(hModuleCurrent, hResDna);
		void* pDna = LockResource(hDna);
		DWORD sizeDna = SizeofResource(hModuleCurrent, hResDna);

		headerLength = min(sizeDna, MAX_HEADER_LENGTH);
		SafeByteArray dnaBytes(pDna, headerLength);

		void* pData = ((LPSAFEARRAY)dnaBytes)->pvData;
		CopyMemory(headerBuffer, pData, headerLength);
	}
	else
	{
		SafeFile dnaFile;
		std::wstring dnaPath(GetAddInFullPath());
		RenameExtension(dnaPath, L".dna");
		if (!FileExists(dnaPath.c_str()))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNANOTFOUND,
					IDS_MSG_BODY_DNAPATHNOTEXIST,
					IDS_MSG_FOOTER_ENSUREDNAFILE,
					hr);
			}
			return E_FAIL;
		}
		hr = dnaFile.Create(dnaPath, GENERIC_READ, FILE_SHARE_READ, OPEN_EXISTING);
		if (FAILED(hr))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNAPROBLEM,
					IDS_MSG_BODY_DNAOPENFAILED,
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
			}
			return E_FAIL;
		}
		hr = dnaFile.Read((LPVOID)headerBuffer, MAX_HEADER_LENGTH, headerLength);
		if (FAILED(hr))
		{
			if (showErrors)
			{
				ShowMessage(IDS_MSG_HEADER_DNAPROBLEM,
					IDS_MSG_BODY_DNAOPENFAILED,
					IDS_MSG_FOOTER_UNEXPECTED,
					hr);
			}
			return E_FAIL;
		}
	}
	if (IsBufferUTF8(headerBuffer, headerLength))
	{
		header = UTF8toUTF16(std::string((char*)headerBuffer, headerLength));
	}
	else
	{
		header = std::wstring((wchar_t*)headerBuffer, headerLength / 2);
	}
	return S_OK;
}

// Returns	S_OK if the attribute was found and read into the attributeValue string.
//			S_FALSE if the attribute was not found at all
//			E_FAIL if there was an XML syntax error in the tag.
// TODO: There is a bug here - I don't check the character before attributeName starts, so I also match XXXName="NotMyName"
//		 For not the .dna schema does not define any conflicts here, but it's not great.
HRESULT GetAttributeValue(std::wstring tag, std::wstring attributeName, std::wstring& attributeValue)
{
	attributeName += L"=";
	size_t attributeNameLength = attributeName.size();

	size_t attributeNameStart = tag.find(attributeName);
	if (attributeNameStart == -1)
	{
		return S_FALSE;
	}

	wchar_t quoteChar = tag[attributeNameStart + attributeNameLength];
	if (quoteChar != L'\'' && quoteChar != L'\"')
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}

	size_t attributeValueStart = attributeNameStart + attributeNameLength + 1;
	size_t attributeValueEnd = tag.find(quoteChar, attributeValueStart);
	if (attributeValueEnd == -1)
	{
		// XML syntax error - not a valid attribute.
		return E_FAIL;
	}
	attributeValue = tag.substr(attributeValueStart, attributeValueEnd - attributeValueStart);
	return S_OK;
}

HRESULT ParseDnaHeader(std::wstring header, std::wstring& addInName, std::wstring& runtimeVersion, bool& shadowCopyFiles, std::wstring& createSandboxedAppDomain, bool& cancelAddInIsolation, bool& disableAssemblyContextUnload, std::wstring& customRuntimeConfiguration, std::wstring& rollForward, std::wstring& runtimeFrameworkVersion)
{
	HRESULT hr;

	size_t rootTagStart = header.find(L"<DnaLibrary");
	if (rootTagStart == -1)
	{
		// Parse error
		return E_FAIL;
	}

	size_t rootTagEnd = header.find(L">", rootTagStart);
	if (rootTagEnd == -1)
	{
		// Parse error
		return E_FAIL;
	}

	std::wstring rootTag = header.substr(rootTagStart, rootTagEnd - rootTagStart + 1);

	// CONSIDER: Some checks, e.g. "v.X..."
	hr = GetAttributeValue(rootTag, L"RuntimeVersion", runtimeVersion);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		runtimeVersion = L"v6.0";
		hr = S_OK;
	}

	std::wstring shadowCopyFilesValue;
	hr = GetAttributeValue(rootTag, L"ShadowCopyFiles", shadowCopyFilesValue);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		shadowCopyFiles = false;
		hr = S_OK;
	}
	else // attribute read OK
	{
		if (CompareNoCase(shadowCopyFilesValue, L"true") == 0)
			shadowCopyFiles = true;
		else
			shadowCopyFiles = false;
	}

	hr = GetAttributeValue(rootTag, L"CreateSandboxedAppDomain", createSandboxedAppDomain);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		createSandboxedAppDomain = L"";
		hr = S_OK;
	}

	std::wstring cancelAddInIsolationValue;
	hr = GetAttributeValue(rootTag, L"Unsafe_CancelAddInIsolation_Unsafe", cancelAddInIsolationValue);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		cancelAddInIsolation = false;
		hr = S_OK;
	}
	else // attribute read OK
	{
		if (CompareNoCase(cancelAddInIsolationValue, L"true") == 0)
			cancelAddInIsolation = true;
		else
			cancelAddInIsolation = false;
	}

	std::wstring disableAssemblyContextUnloadValue;
	hr = GetAttributeValue(rootTag, L"DisableAssemblyContextUnload", disableAssemblyContextUnloadValue);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		disableAssemblyContextUnload = false;
		hr = S_OK;
	}
	else // attribute read OK
	{
		if (CompareNoCase(disableAssemblyContextUnloadValue, L"true") == 0)
			disableAssemblyContextUnload = true;
		else
			disableAssemblyContextUnload = false;
	}

	hr = GetAttributeValue(rootTag, L"CustomRuntimeConfiguration", customRuntimeConfiguration);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		customRuntimeConfiguration = L"";
		hr = S_OK;
	}

	hr = GetAttributeValue(rootTag, L"Name", addInName);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		addInName = L"";
		hr = S_OK;
	}

	hr = GetAttributeValue(rootTag, L"RollForward", rollForward);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		rollForward = L"";
		hr = S_OK;
	}

	hr = GetAttributeValue(rootTag, L"RuntimeFrameworkVersion", runtimeFrameworkVersion);
	if (FAILED(hr))
	{
		// Parse error
		return E_FAIL;
	}
	if (hr == S_FALSE)
	{
		runtimeFrameworkVersion = L"";
		hr = S_OK;
	}

	return hr;
}

HRESULT GetAddInName(std::wstring& addInName)
{
	HRESULT hr;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	bool disableAssemblyContextUnload;
	std::wstring customRuntimeConfiguration;
	std::wstring rollForward;
	std::wstring runtimeFrameworkVersion;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
		if (addInName.empty())
		{
			std::wstring xllPath(GetAddInFullPath());
			StripPath(xllPath);
			RemoveExtension(xllPath);
			addInName = xllPath;
		}
	}
	return hr;
}

HRESULT GetDisableAssemblyContextUnload(bool& disableAssemblyContextUnload)
{
	HRESULT hr;
	std::wstring addInName;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	std::wstring customRuntimeConfiguration;
	std::wstring rollForward;
	std::wstring runtimeFrameworkVersion;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
	}
	return hr;
}

HRESULT GetCustomRuntimeConfiguration(std::wstring& customRuntimeConfiguration)
{
	HRESULT hr;
	std::wstring addInName;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	bool disableAssemblyContextUnload;
	std::wstring rollForward;
	std::wstring runtimeFrameworkVersion;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
	}
	return hr;
}

HRESULT GetMajorRuntimeVersion(int& majorRuntimeVersion)
{
	HRESULT hr;
	std::wstring addInName;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	bool disableAssemblyContextUnload;
	std::wstring customRuntimeConfiguration;
	std::wstring rollForward;
	std::wstring runtimeFrameworkVersion;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr) || clrVersion.length() < 2 || clrVersion[0] != L'v')
		{
			return E_FAIL;
		}

		try
		{
			majorRuntimeVersion = std::stoi(clrVersion.substr(1));
		}
		catch (const std::exception&)
		{
			return E_FAIL;
		}
	}
	return hr;
}

HRESULT GetRollForward(std::wstring& rollForward)
{
	HRESULT hr;
	std::wstring addInName;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	bool disableAssemblyContextUnload;
	std::wstring customRuntimeConfiguration;
	std::wstring runtimeFrameworkVersion;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
	}
	return hr;
}

HRESULT GetRuntimeFrameworkVersion(std::wstring& runtimeFrameworkVersion)
{
	HRESULT hr;
	std::wstring addInName;
	std::wstring header;
	std::wstring clrVersion;
	bool shadowCopyFiles;
	std::wstring createSandboxedAppDomainValue;
	bool cancelAddInIsolation;
	bool disableAssemblyContextUnload;
	std::wstring customRuntimeConfiguration;
	std::wstring rollForward;
	hr = GetDnaHeader(false, header);	// Don't show errors here.
	if (!FAILED(hr))
	{
		hr = ParseDnaHeader(header, addInName, clrVersion, shadowCopyFiles, createSandboxedAppDomainValue, cancelAddInIsolation, disableAssemblyContextUnload, customRuntimeConfiguration, rollForward, runtimeFrameworkVersion); // No errors yet.
		if (FAILED(hr))
		{
			return E_FAIL;
		}
	}
	return hr;
}
