#pragma once
#include <string>

// Creates a random dir in %temp%\topDirName.
// In the destructor tries to delete this dir. If fails due to locked files, creates "pending delete" file.
// On startup tries to delete all dirs in %temp%\topDirName with the "pending delete" file.

class TempDir
{
public:
	TempDir(const std::wstring& topDirName);
	~TempDir();
	const std::wstring& GetPath() const;

private:
	void DeletePendingDirs();
	void DeleteDir(const std::wstring& dirPath, bool createPendingDeleteFile);

	std::wstring path, topDirPath;
};
