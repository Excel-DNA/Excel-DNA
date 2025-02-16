//  Copyright (c) Govert van Drimmelen. All rights reserved.

#pragma once

#include <Windows.h>
#include <string>

HRESULT GetAddInName(std::wstring& addInName);
HRESULT GetDisableAssemblyContextUnload(bool& disableAssemblyContextUnload);
HRESULT GetCustomRuntimeConfiguration(std::wstring& customRuntimeConfiguration);
HRESULT GetMajorRuntimeVersion(int& majorRuntimeVersion);
HRESULT GetRollForward(std::wstring& rollForward);
HRESULT GetRuntimeFrameworkVersion(std::wstring& runtimeFrameworkVersion);
