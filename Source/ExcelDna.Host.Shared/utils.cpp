// This code was contributed by Koustubh Moharir (Thank you!)

#include "utils.h"
#include <windows.h>
#include <tchar.h>
#include <comdef.h>
#include <algorithm>
#include <iostream>
#include <filesystem>
#include "dnainfo.h"

extern HMODULE hModuleCurrent;

TempFileHolder::~TempFileHolder()
{
	destroy();
}
void TempFileHolder::setFileName(std::wstring fileName)
{
	this->fileName = fileName;
}
void TempFileHolder::destroy()
{
	if (!this->fileName.empty())
	{
		::DeleteFile(fileName.c_str());
		this->fileName = L"";
	}
}

SafeVariantArray::SafeVariantArray(int size)
{
	pArray = SafeArrayCreateVector(VT_VARIANT, 0, size);
}

SafeVariantArray::~SafeVariantArray()
{
	SafeArrayDestroy(pArray);
}

void SafeVariantArray::lock()
{
	SafeArrayLock(pArray);
}
void SafeVariantArray::unlock()
{
	SafeArrayUnlock(pArray);
}
void SafeVariantArray::setElement(int i, INT32 value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_I4;
	element.lVal = value;
}
void SafeVariantArray::setElement(int i, INT64 value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_I8;
	element.llVal = value;
}
void SafeVariantArray::setElement(int i, BSTR value)
{
	auto& element = static_cast<VARIANT*>(pArray->pvData)[i];
	VariantInit(&element);
	element.vt = VT_BSTR;
	element.bstrVal = value;
}

SafeByteArray::SafeByteArray(void* data, int sizeInBytes)
{
	pArray = SafeArrayCreateVector(VT_UI1, 0, sizeInBytes);
	SafeArrayLock(pArray);
	memcpy(pArray->pvData, data, sizeInBytes);
	SafeArrayUnlock(pArray);
}
SafeByteArray::~SafeByteArray()
{
	SafeArrayDestroy(pArray);
}

int SafeByteArray::AccessData(byte** ppData)
{
	SafeArrayAccessData(pArray, (void**)ppData);
	return pArray->rgsabound->cElements;
}

void SafeByteArray::UnaccessData()
{
	SafeArrayUnaccessData(pArray);
}

bool FileExists(LPCTSTR szPath)
{
	DWORD dwAttrib = GetFileAttributes(szPath);

	return (dwAttrib != INVALID_FILE_ATTRIBUTES &&
		!(dwAttrib & FILE_ATTRIBUTE_DIRECTORY));
}

std::wstring LoadStringFromResource(HMODULE hModule, int id)
{
	const wchar_t* buffer = nullptr;
	LoadStringW(hModule, id, (LPWSTR)&buffer, 0);//ugly cast for badly designed API (specifying buffer size == 0 returns a read only pointer.
	return std::wstring(buffer, *((WORD*)buffer - 1));//The WORD preceding the address is the size of the resource string. which is not null-terminated.
}

std::wstring FormatString(std::wstring formatString, ...)
{
	LPWSTR buffer = NULL;
	va_list args = NULL;
	va_start(args, formatString);
	FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER,
		formatString.c_str(),
		0,
		0,
		(LPWSTR)&buffer,//ugly cast for badly designed API (FORMAT_MESSAGE_ALLOCATE_BUFFER returns a system allocated buffer)
		0,
		&args);

	va_end(args);

	std::wstring result(buffer ? buffer : L"");
	if (buffer)
		LocalFree(buffer);
	return result;
}

void RemoveFileSpecFromPath(std::wstring& filePath)
{
	size_t dirSepInd = filePath.find_last_of(L'\\');
	if (dirSepInd >= 0)
		filePath.erase(dirSepInd);
}

void StripPath(std::wstring& filePath)
{
	size_t dirSepInd = filePath.find_last_of(L'\\');
	if (dirSepInd >= 0)
		filePath.erase(0, dirSepInd + 1);
}

std::wstring GetDirectory(const std::wstring& filePath)
{
	std::wstring result(filePath);
	RemoveFileSpecFromPath(result);
	return result;
}

std::wstring GetDirectoryName(const std::wstring& filePath)
{
	std::wstring result(GetDirectory(filePath));
	StripPath(result);
	return result;
}

void RemoveExtension(std::wstring& filePath)
{
	size_t dotInd = filePath.find_last_of(L'.');
	filePath.erase(dotInd);
}

void RenameExtension(std::wstring& filePath, std::wstring ext)
{
	size_t dotInd = filePath.find_last_of(L'.');
	filePath.replace(dotInd, std::wstring::npos, ext);
}

HRESULT HResultFromLastError()
{
	return HRESULT_FROM_WIN32(GetLastError());
}

std::wstring GetLastErrorMessage()
{
	return _com_error(HResultFromLastError()).ErrorMessage();
}

SafeFile::SafeFile()
	: handle(INVALID_HANDLE_VALUE)
{

}

SafeFile::~SafeFile()
{
	if (handle != INVALID_HANDLE_VALUE)
		CloseHandle(handle);
}

HRESULT SafeFile::Create(std::wstring path, DWORD dwDesiredAccess, DWORD dwShareMode, DWORD dwCreationDisposition)
{
	handle = ::CreateFileW(path.c_str(), dwDesiredAccess, dwShareMode,
		NULL, dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, NULL);

	if (handle == INVALID_HANDLE_VALUE)
		return HResultFromLastError();
	return S_OK;
}

HRESULT SafeFile::Read(LPVOID pBuffer, DWORD nBufSize, DWORD& nBytesRead)
{
	BOOL b = ::ReadFile(handle, pBuffer, nBufSize, &nBytesRead, NULL);
	if (!b)
		return HResultFromLastError();

	return S_OK;
}

std::wstring UTF8toUTF16(const std::string& utf8)
{
	std::wstring utf16;
	int len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, NULL, 0);
	if (len > 1)
	{
		utf16.resize(len);
		len = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, &utf16[0], len);
		utf16.resize(len);
	}
	return utf16;
}

std::string ANSIWStringToString(const std::wstring& ws)
{
	std::string result(ws.length(), 0);
	std::transform(ws.begin(), ws.end(), result.begin(), [](wchar_t c) {
		return static_cast<char>(c);
		});
	return result;
}

HRESULT WriteAllBytes(const std::wstring& filePath, void* buf, DWORD size)
{
	HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
		return HResultFromLastError();

	DWORD dwBytesWritten;
	if (!WriteFile(hFile, buf, size, &dwBytesWritten, NULL))
	{
		HRESULT hr = HResultFromLastError();
		CloseHandle(hFile);
		return hr;
	}

	if (!CloseHandle(hFile))
		return HResultFromLastError();

	return S_OK;
}

std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2)
{
	return path1 + L"\\" + path2;
}

std::wstring PathCombine(const std::wstring& path1, const std::wstring& path2, const std::wstring& path3)
{
	return PathCombine(PathCombine(path1, path2), path3);
}

BOOL IsRunningOnCluster()
{
	// Our check is to see if the current process is called Excel.exe.
	// Hopefully this doen't change soon.

	TCHAR hostPathName[MAX_PATH];
	DWORD count = GetModuleFileName(NULL, hostPathName, MAX_PATH);

	std::wstring hostPath = hostPathName;
	StripPath(hostPath);

	if (CompareNoCase(hostPath, L"EXCEL.EXE") == 0)
	{
		return false;
	}

	return true;
}

struct FindExcelWindowParam
{
	DWORD processId;
	HWND  hwndFound;
};

BOOL WINAPI FindExcelWindowCallback(HWND hwnd, LPARAM lParam)
{
	FindExcelWindowParam* pParam = (FindExcelWindowParam*)lParam;
	DWORD processId = 0;
	GetWindowThreadProcessId(hwnd, &processId);
	if (processId == pParam->processId)
	{
		wchar_t className[11];
		DWORD count = RealGetWindowClass(hwnd, className, 10);
		if (_tcsncmp(className, L"XLMAIN", 6))
		{
			pParam->hwndFound = hwnd;
			SetLastError(0);
			return FALSE;
		}
	}
	return TRUE;
}

HWND FindCurrentExcelWindow()
{
	FindExcelWindowParam param;
	param.processId = GetCurrentProcessId();
	param.hwndFound = NULL;

	EnumWindows(FindExcelWindowCallback, (LPARAM)&param);
	return param.hwndFound;
}

void ShowMessageError(HWND hwndParent)
{
	MessageBox(hwndParent, L"There was problem while loading the add-in. \r\nA detailed message could not be displayed.", L"Add-In Loader", MB_ICONEXCLAMATION);
}

void ShowMessage(int headerId, int bodyId, int footerId, HRESULT hr)
{
	if (IsRunningOnCluster())
	{
		// TODO: Consider what to do in cluster context?
		return;
	}

	HWND hwndExcel = FindCurrentExcelWindow();
	try
	{
		std::wstring  addInFullPath = GetAddInFullPath();
		std::wstring  addInFileName = addInFullPath;
		StripPath(addInFileName);

		std::wstring msgTitle = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_TITLE), addInFileName.c_str());

		std::wstring header = LoadStringFromResource(hModuleCurrent, headerId);
		std::wstring body = LoadStringFromResource(hModuleCurrent, bodyId);
		std::wstring footer = LoadStringFromResource(hModuleCurrent, footerId);

		std::wstring hresult = L"";
		if (hr != S_OK)
		{
			_com_error error(hr);
			hresult = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_HRESULT), error.ErrorMessage());
		}

		std::wstring msg = FormatString(LoadStringFromResource(hModuleCurrent, IDS_MSG_TEMPLATE), header.c_str(), body.c_str(), footer.c_str(), hresult.c_str(), addInFullPath.c_str());
		MessageBox(hwndExcel, msg.c_str(), msgTitle.c_str(), MB_ICONEXCLAMATION);
	}
	catch (...)
	{
		ShowMessageError(hwndExcel);
	}
}

void ShowHostError(const std::wstring& msg)
{
	std::wstring title = L"ExcelDna.Host";

	std::wstring addInName;
	HRESULT hr = GetAddInName(addInName);
	if (!FAILED(hr))
	{
		title = addInName + L" - " + title;
	}

	MessageBox(FindCurrentExcelWindow(), msg.c_str(), title.c_str(), MB_OK | MB_ICONERROR);
}

std::wstring GetAddInFullPath()
{
	wchar_t buffer[MAX_PATH];
	DWORD count = GetModuleFileName(hModuleCurrent, buffer, MAX_PATH);
	return std::wstring(buffer);
}

BOOL IsBufferUTF8(BYTE* buffer, DWORD bufferLength)
{
	// Only UTF-8 and UTF-16 is supported (here)
	// The check here is naive - does not read the xml processing instruction.
	// CONSIDER: Use WIN32 API function IsTextUnicode ?

	// Check for byte order marks.
	if (bufferLength < 3)
	{
		// Doesn't matter - will fail later.
		return true;
	}
	if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
	{
		// Standard UTF-8 BOM
		return true;
	}
	//if (buffer[0] == 0xFF && buffer[1] == 0xFE && buffer[2] == 0x00 && buffer[3] == 0x00)
	//{
	//	// UTF-32 LE
	//	return false;
	//}
	//if (buffer[0] == 0x00 && buffer[1] == 0x00 && buffer[2] == 0xFE && buffer[3] == 0xFF)
	//{
	//	// UTF-32 BE
	//	return false;
	//}
	if (buffer[0] == 0xFF && buffer[1] == 0xFE)
	{
		// UTF-16 LE
		return false;
	}
	if (buffer[0] == 0xFE && buffer[1] == 0xFF)
	{
		// UTF-16 BE
		return false;
	}
	// Might be ANSI or some other code page. Treated as UTF-8 here.
	return true;
}

int WriteResourceToFile(HMODULE hModuleXll, const std::wstring& resourceName, const std::wstring& resourceType, const std::wstring& filePath)
{
	HRSRC hResManagedHost = FindResource(hModuleXll, resourceName.c_str(), resourceType.c_str());
	if (hResManagedHost == NULL)
	{
		ShowHostError(L"Failure to find resource " + resourceName);
		return EXIT_FAILURE;
	}

	HGLOBAL hManagedHost = LoadResource(hModuleXll, hResManagedHost);
	if (hManagedHost == NULL)
	{
		ShowHostError(L"Failure to load resource " + resourceName);
		return EXIT_FAILURE;
	}

	void* buf = LockResource(hManagedHost);
	if (buf == NULL)
	{
		ShowHostError(L"Failure to lock resource " + resourceName);
		return EXIT_FAILURE;
	}

	DWORD resSize = SizeofResource(hModuleXll, hResManagedHost);
	SafeByteArray safeBytes(buf, resSize);
	byte* pData;
	int nSize = safeBytes.AccessData(&pData);

	HRESULT hr = WriteAllBytes(filePath, pData, nSize);
	if (FAILED(hr))
	{
		std::wstringstream stream;
		stream << L"Saving " << resourceName << L" failed: " << std::hex << std::showbase << hr;
		ShowHostError(stream.str());
		return EXIT_FAILURE;
	}

	return EXIT_SUCCESS;
}
