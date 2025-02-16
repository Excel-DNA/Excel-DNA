#include "TempDir.h"
#include "utils.h"

#include <shlwapi.h>
#undef PathCombine
#pragma comment(lib, "Shlwapi.lib")

#include <vector>
#include <algorithm>
#include <functional>
#include <cassert>
#include <rpc.h>
#pragma comment(lib, "rpcrt4.lib")

namespace
{
	bool IsSpecialDir(const std::wstring& dirName)
	{
		std::vector<std::wstring> specialDirs = { L".", L".." };
		return std::any_of(specialDirs.begin(), specialDirs.end(), [&dirName](auto i) { return i == dirName; });
	}

	std::wstring CreateUuid()
	{
		UUID uuid;
		RPC_STATUS s = UuidCreate(&uuid);
		assert((s == RPC_S_OK || s == RPC_S_UUID_LOCAL_ONLY) && "Failure: UuidCreate");

		RPC_WSTR str;
		s = UuidToString(&uuid, &str);
		assert(s == RPC_S_OK && "Failure: UuidToString");

		std::wstring result((TCHAR*)str);
		RpcStringFree(&str);
		return result;
	}

	constexpr auto pendingDeleteFileName(L"pending delete");
}

TempDir::TempDir(const std::wstring& topDirName)
{
	TCHAR lpTempPathBuffer[MAX_PATH];
	DWORD dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);
	assert(!(dwRetVal > MAX_PATH || dwRetVal == 0) && "Failure: GetTempPath");

	topDirPath = PathCombine(lpTempPathBuffer, topDirName);
	CreateDirectory(topDirPath.c_str(), NULL);

	DeletePendingDirs();

	path = PathCombine(topDirPath, CreateUuid());
	BOOL b = CreateDirectory(path.c_str(), NULL);
	assert(b && "Failure: CreateDirectory");
}

TempDir::~TempDir()
{
	DeleteDir(path, true);
}

const std::wstring& TempDir::GetPath() const
{
	return path;
}

void TempDir::DeletePendingDirs()
{
	WIN32_FIND_DATA ffd;
	HANDLE hFind = FindFirstFile(PathCombine(topDirPath, L"*").c_str(), &ffd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY && !IsSpecialDir(ffd.cFileName))
			{
				std::wstring subdir = PathCombine(topDirPath, ffd.cFileName);
				if (PathFileExists(PathCombine(subdir, pendingDeleteFileName).c_str()))
					DeleteDir(subdir, true);
			}
		} while (FindNextFile(hFind, &ffd) != 0);
		FindClose(hFind);
	}
}

void TempDir::DeleteDir(const std::wstring& dirPath, bool createPendingDeleteFile)
{
	WIN32_FIND_DATA ffd;
	HANDLE hFind = FindFirstFile(PathCombine(dirPath, L"*").c_str(), &ffd);
	if (hFind != INVALID_HANDLE_VALUE)
	{
		do
		{
			if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0)
				DeleteFile(PathCombine(dirPath, ffd.cFileName).c_str());
			else if ((ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == FILE_ATTRIBUTE_DIRECTORY && !IsSpecialDir(ffd.cFileName))
				DeleteDir(PathCombine(dirPath, ffd.cFileName), false);
		} while (FindNextFile(hFind, &ffd) != 0);
		FindClose(hFind);
	}
	if (RemoveDirectory(dirPath.c_str()) == 0)
	{
		if (createPendingDeleteFile)
		{
			HANDLE hFile = CreateFile(PathCombine(dirPath, pendingDeleteFileName).c_str(), GENERIC_WRITE, 0, NULL, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
			if (hFile != INVALID_HANDLE_VALUE)
				CloseHandle(hFile);
		}
	}
}
