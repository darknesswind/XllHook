#include <ctype.h>
#include <windows.h>
#include "xlcall.h"
#include <framewrk.h>
#include "loghelper.h"
#include "ExcelProcxy.h"
#include "xllhook.h"
#define __CreateMenu 1

static const UINT uRegFuncCount = 5;
static LPWSTR rgFuncs[uRegFuncCount][7] = {
	{ L"XllClearLog",	L"I",	L"XllClearLog" },
	{ L"XllOpenFolder",	L"I",	L"XllOpenFolder" },
	{ L"XllPauseLog",	L"I",	L"XllPauseLog" },
	{ L"XllResumeLog",	L"I",	L"XllResumeLog" },
	{ L"XllTest",		L"JJ",	L"XllTest" },
};

static LPSTR rgFuncs4[uRegFuncCount][7] = {
	{ "XllClearLog",	"I",	"XllClearLog" },
	{ "XllOpenFolder",	"I",	"XllOpenFolder" },
	{ "XllPauseLog",	"I",	"XllPauseLog" },
	{ "XllResumeLog",	"I",	"XllResumeLog" },
	{ "XllTest",		"JJ",	"XllTest" },
};

#define g_rgMenuRows 5
#define g_rgMenuCols 5

static LPWSTR g_rgMenu12[g_rgMenuRows][g_rgMenuCols] =
{
	{ L"&XllHook", L"", L"",
	L"XllHook Add-In", L"" },
	{ L"&ClearLog", L"XllClearLog", L"",
	L"ClearCurLog", L"" },
	{ L"&OpenLogFolder", L"XllOpenFolder", L"",
	L"Open Current Log Folder", L"" },
	{ L"&PauseLog", L"XllPauseLog", L"",
	L"&Pause Log", L"" },
	{ L"&ResumeLog", L"XllResumeLog", L"",
	L"Resume Log", L"" },
};

static LPSTR g_rgMenu4[g_rgMenuRows][g_rgMenuCols] =
{
	{	"&XllHook",		"",					"",	"XllHook Add-In",	"" },
	{	"&ClearLog",	"XllClearLog",		"",	"ClearCurLog",		"" },
	{	"&OpenFolder",	"XllOpenFolder",	"",	"Open Current Log Folder", "" },
	{	"&PauseLog",	"XllPauseLog",		"",	"&Pause Log",		"" },
	{	"&ResumeLog",	"XllResumeLog",		"",	"&Resume Log",		"" },
};

// 不注册函数的话Excel会自动把XLL卸载掉
__declspec(dllexport) short WINAPI XllClearLog(void)
{
	LogHelper::Instance().ClearLog();
	return 1;
}

__declspec(dllexport) short WINAPI XllOpenFolder(void)
{
	LogHelper::Instance().OpenFolder();
	return 1;
}

__declspec(dllexport) short WINAPI XllPauseLog(void)
{
	LogHelper::Instance().PauseLog();
	return 1;
}

__declspec(dllexport) short WINAPI XllResumeLog(void)
{
	LogHelper::Instance().ResumeLog();
	return 1;
}

__declspec(dllexport) int WINAPI XllTest(int val)
{
	return val;
}

void Excel4AutoOpen();
void Excel12AutoOpen();
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
	LogHelper::Instance().PauseLog();

	UINT uExcelVersion = (XLCallVer() >> 8);
	if (uExcelVersion < 12)	// Version < Excel 2007
		Excel4AutoOpen();
	else
		Excel12AutoOpen();

	LogHelper::Instance().SetOpened(true);
	LogHelper::Instance().ResumeLog();
	return 1;
}
void Excel4AutoOpen()
{
	static XLOPER xDLL, xTest, xMenu, xResult;
	Excel4(xlcall::xlGetName, &xDLL, 0);

	for (int i = 0; i < uRegFuncCount; i++)
	{
		Excel4(xlcall::xlfRegister, &xResult, 4,
			(LPXLOPER)&xDLL,
			(LPXLOPER)TempStrConst(rgFuncs4[i][0]),
			(LPXLOPER)TempStrConst(rgFuncs4[i][1]),
			(LPXLOPER)TempStrConst(rgFuncs4[i][2]));
	}
	/* Free the XLL filename */
	Excel4(xlcall::xlFree, 0, 1, (LPXLOPER)&xDLL);

#if __CreateMenu
	Excel4(xlcall::xlfGetBar, &xTest, 3, TempInt(10), TempStrConst("XllHook"), TempInt(0));
	if (xTest.xltype == xltypeErr)
	{
		HANDLE   hMenu;		   // global memory holding menu //
		LPXLOPER pxMenu;	   // Points to first menu item //
		LPXLOPER px;		   // Points to the current item //
		hMenu = GlobalAlloc(GMEM_MOVEABLE, sizeof(XLOPER) * g_rgMenuCols * g_rgMenuRows);
		px = pxMenu = (LPXLOPER)GlobalLock(hMenu);

		for (int i = 0; i < g_rgMenuRows; i++)
		{
			for (int j = 0; j < g_rgMenuCols; j++)
			{
				px->xltype = xltypeStr;
				px->val.str = TempStrConst(g_rgMenu4[i][j])->val.str;
				px++;
			}
		}

		xMenu.xltype = xltypeMulti;
		xMenu.val.array.lparray = pxMenu;
		xMenu.val.array.rows = g_rgMenuRows;
		xMenu.val.array.columns = g_rgMenuCols;

		Excel4(xlcall::xlfAddMenu, 0, 3, TempNum(10), (LPXLOPER)&xMenu, TempStrConst("Help"));

		GlobalUnlock(hMenu);
		GlobalFree(hMenu);
	}
	Excel4(xlcall::xlFree, 0, 1, (LPXLOPER)&xTest);
#endif
}

void Excel12AutoOpen()
{
	static XLOPER12 xDLL, xTest, xMenu;
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

#if __CreateMenu
	Excel12f(xlcall::xlfGetBar, &xTest, 3, TempInt12(10), TempStr12(L"XllHook"), TempInt12(0));
	if (xTest.xltype == xltypeErr)
	{
		HANDLE   hMenu;		   // global memory holding menu //
		LPXLOPER12 pxMenu;	   // Points to first menu item //
		LPXLOPER12 px;		   // Points to the current item //
		hMenu = GlobalAlloc(GMEM_MOVEABLE, sizeof(XLOPER12) * g_rgMenuCols * g_rgMenuRows);
		px = pxMenu = (LPXLOPER12)GlobalLock(hMenu);

		for (int i = 0; i < g_rgMenuRows; i++)
		{
			for (int j = 0; j < g_rgMenuCols; j++)
			{
				px->xltype = xltypeStr;
				px->val.str = TempStr12(g_rgMenu12[i][j])->val.str;
				px++;
			}
		}

		xMenu.xltype = xltypeMulti;
		xMenu.val.array.lparray = pxMenu;
		xMenu.val.array.rows = g_rgMenuRows;
		xMenu.val.array.columns = g_rgMenuCols;

		Excel12f(xlcall::xlfAddMenu, 0, 3, TempNum12(10), (LPXLOPER12)&xMenu, TempStr12(L"Help"));

		GlobalUnlock(hMenu);
		GlobalFree(hMenu);
	}
	Excel12f(xlcall::xlFree, 0, 1, (LPXLOPER12)&xTest);
#endif
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
	LogHelper::Instance().PauseLog();

	UINT uExcelVersion = (XLCallVer() >> 8);
	if (uExcelVersion < 12)	// Version < Excel 2007
	{
		static XLOPER xRes;
		Excel4(xlcall::xlfGetBar, &xRes, 3, TempInt(10), TempStrConst("Generic"), TempInt(0));
		if (xRes.xltype != xltypeErr)
		{
			Excel4(xlcall::xlfDeleteMenu, 0, 2, TempNum(10), TempStrConst("Generic"));
			// Free the XLOPER12 returned by xlfGetBar //
			Excel4(xlcall::xlFree, 0, 1, (LPXLOPER)&xRes);
		}
	}
	else
	{
		static XLOPER12 xRes;
		Excel12f(xlcall::xlfGetBar, &xRes, 3, TempInt12(10), TempStr12(L"Generic"), TempInt12(0));
		if (xRes.xltype != xltypeErr)
		{
			Excel12f(xlcall::xlfDeleteMenu, 0, 2, TempNum12(10), TempStr12(L"Generic"));
			// Free the XLOPER12 returned by xlfGetBar //
			Excel12f(xlcall::xlFree, 0, 1, (LPXLOPER12)&xRes);
		}
	}
	LogHelper::Instance().SetOpened(false);
	LogHelper::Instance().ResumeLog();
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
		xInfo.val.str = "\014!!!!!DBGHOOK";
// 		xInfo.val.str = "\tFixWndTxt";
#else
		xInfo.val.str = "\014!!!!!XLLHOOK";
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
#ifdef _DEBUG
#	if SET_Hook_XLLExport
		xInfo.val.str = L"\014!FullDbgHook";
#	elif SET_Hook_XLL
		xInfo.val.str = L"\014!!!!!XLLHOOK";
#	else
		xInfo.val.str = L"\014!!!OTHERHOOK";
#	endif
#else
		xInfo.val.str = L"\014!!!!!XLLHOOK";
#endif
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

