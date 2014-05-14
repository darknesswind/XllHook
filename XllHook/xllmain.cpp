#include <ctype.h>
#include <windows.h>
#include "xlcall.h"
#include <framewrk.h>
#include "loghelper.h"

static LPWSTR rgFuncs[1][7] = {
	{ L"XllHookDummy", L"I", L"XllHookDummy" },
};

// 不注册函数的话Excel会自动把XLL卸载掉
__declspec(dllexport) short WINAPI XllHookDummy(void)
{
	return 1;
}

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
	static XLOPER12 xDLL;

	LogHelper::Instance().PauseLog();
	Excel12f(xlcall::xlGetName, &xDLL, 0);

	for (int i = 0; i < 1; i++)
	{
		Excel12f(xlcall::xlfRegister, 0, 4,
			(LPXLOPER12)&xDLL,
			(LPXLOPER12)TempStr12(rgFuncs[i][0]),
			(LPXLOPER12)TempStr12(rgFuncs[i][1]),
			(LPXLOPER12)TempStr12(rgFuncs[i][2]));
	}

	/* Free the XLL filename */
	Excel12f(xlcall::xlFree, 0, 1, (LPXLOPER12)&xDLL);

	LogHelper::Instance().ResumeLog();
	return 1;
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
	return 1;
}

__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xRegId;
	return (LPXLOPER12)&xRegId;
}

__declspec(dllexport) int WINAPI xlAutoAdd(void)
{
	return 1;
}

__declspec(dllexport) int WINAPI xlAutoRemove(void)
{
	return 1;
}

__declspec(dllexport) LPXLOPER12 WINAPI xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction;

	/*
	** This code coerces the passed-in value to an integer. This is how the
	** code determines what is being requested. If it receives a 1, it returns a
	** string representing the long name. If it receives anything else, it
	** returns a #VALUE! error.
	*/
	LogHelper::Instance().PauseLog();
	Excel12f(xlcall::xlCoerce, &xIntAction, 2, xAction, TempInt12(xltypeInt));

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = xltypeStr;
		xInfo.val.str = L"\007!!!!!XLLHOOK";
	}
	else
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms

	LogHelper::Instance().ResumeLog();
	return (LPXLOPER12)&xInfo;
}

