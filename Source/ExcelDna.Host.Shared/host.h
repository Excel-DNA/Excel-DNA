//  Copyright (c) Govert van Drimmelen. All rights reserved.

#pragma once

#include "exports.h"

int load_and_run(const std::wstring& basePath, XlAddInExportInfo* pExportInfo, HMODULE hModuleXll, LPCWSTR pathXll);
