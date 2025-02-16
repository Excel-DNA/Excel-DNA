#include "LPenHelper.h"
#include <Windows.h>

namespace
{
	typedef long(__stdcall* LPenHelperT)(int wCode, void* lpv);

	class PenHelper
	{
	public:
		PenHelper() : proc(NULL)
		{
			handle = LoadLibrary(L"XLCALL32.DLL");
			if (handle != NULL)
				proc = (LPenHelperT)GetProcAddress(handle, "LPenHelper");
		}

		long Invoke(int wCode, void* lpv)
		{
			if (proc != NULL)
			{
				try
				{
					return proc(wCode, lpv);
				}
				catch (...)
				{
				}
			}

			return -1;
		}

		~PenHelper()
		{
			if (handle != NULL)
				FreeLibrary(handle);
		}

	private:
		HINSTANCE handle;
		LPenHelperT proc;
	};

	PenHelper helper;
}

long __stdcall LPenHelper(int wCode, void* lpv)
{
	return helper.Invoke(wCode, lpv);
}
