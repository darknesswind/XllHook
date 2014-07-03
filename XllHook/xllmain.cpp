#include <ctype.h>
#include <windows.h>
#include "xlcall.h"
#include <framewrk.h>
#include "loghelper.h"

static const UINT uRegFuncCount = 1;
static LPWSTR rgFuncs[uRegFuncCount][7] = {
	{ L"XllHookDummy", L"I", L"XllHookDummy" },
};

static LPSTR rgFuncs4[uRegFuncCount][7] = {
	{ "XllHookDummy", "I", "XllHookDummy" },
};

// 不注册函数的话Excel会自动把XLL卸载掉
__declspec(dllexport) short WINAPI XllHookDummy(void)
{
	return 1;
}

__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
	LogHelper::Instance().PauseLog();

	UINT uExcelVersion = (XLCallVer() >> 8);
	if (uExcelVersion < 12)	// Version < Excel 2007
	{
		static XLOPER xDLL;
// 		TempStr(0);
		Excel4(xlcall::xlGetName, &xDLL, 0);

		for (int i = 0; i < uRegFuncCount; i++)
		{
			Excel4(xlcall::xlfRegister, 0, 4,
				(LPXLOPER)&xDLL,
				(LPXLOPER)TempStrConst(rgFuncs4[i][0]),
				(LPXLOPER)TempStrConst(rgFuncs4[i][1]),
				(LPXLOPER)TempStrConst(rgFuncs4[i][2]));
		}

		/* Free the XLL filename */
		Excel4(xlcall::xlFree, 0, 1, (LPXLOPER)&xDLL);
	}
	else
	{
		static XLOPER12 xDLL;

		Excel12f(xlcall::xlGetName, &xDLL, 0);

		for (int i = 0; i < uRegFuncCount; i++)
		{
			Excel12f(xlcall::xlfRegister, 0, 4,
				(LPXLOPER12)&xDLL,
				(LPXLOPER12)TempStr12(rgFuncs[i][0]),
				(LPXLOPER12)TempStr12(rgFuncs[i][1]),
				(LPXLOPER12)TempStr12(rgFuncs[i][2]));
		}

		/* Free the XLL filename */
		Excel12f(xlcall::xlFree, 0, 1, (LPXLOPER12)&xDLL);
	}

	LogHelper::Instance().ResumeLog();
	return 1;
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
	return 1;
}

__declspec(dllexport) LPXLOPER WINAPI xlAutoRegister(LPXLOPER pxName)
{
	static XLOPER xRegId;
	xRegId.xltype = xltypeMissing;
	return (LPXLOPER)&xRegId;
}

__declspec(dllexport) LPXLOPER12 WINAPI xlAutoRegister12(LPXLOPER12 pxName)
{
	static XLOPER12 xRegId;
	xRegId.xltype = xltypeMissing;
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

__declspec(dllexport) LPXLOPER WINAPI xlAddInManagerInfo(LPXLOPER xAction)
{
	static XLOPER xInfo, xIntAction;

	/*
	** This code coerces the passed-in value to an integer. This is how the
	** code determines what is being requested. If it receives a 1, it returns a
	** string representing the long name. If it receives anything else, it
	** returns a #VALUE! error.
	*/
	LogHelper::Instance().PauseLog();
	Excel4(xlcall::xlCoerce, &xIntAction, 2, xAction, TempInt(xltypeInt));

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = xltypeStr;
#ifdef _DEBUG
		xInfo.val.str = "\007!!!!!DBGHOOK";
#else
		xInfo.val.str = "\007!!!!!XLLHOOK";
#endif
	}
	else
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPERs is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms

	LogHelper::Instance().ResumeLog();
	return (LPXLOPER)&xInfo;
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

