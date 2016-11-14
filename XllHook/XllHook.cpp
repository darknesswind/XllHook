#include <ctype.h>
#include <windows.h>
#include <WinGDI.h>
#include "detver.h"
#include "detours.h"
#include "syelog.h"
#include <cstdio>
#include <cassert>
#include "ExcelProcxy.h"
#include "loghelper.h"
#include "xlcall.h"
#include "XllHook.h"
#include <map>

BOOL ProcessAttach(HMODULE hDll);
BOOL ProcessDetach(HMODULE hDll);
LONG AttachDetours(VOID);
LONG DetachDetours(VOID);
BOOL ThreadAttach(HMODULE hDll);
BOOL ThreadDetach(HMODULE hDll);

static HMODULE s_hExcel = NULL;
static HMODULE s_hInst = NULL;
static HMODULE s_hThis = NULL;
static WCHAR s_wzDllPath[MAX_PATH + 1];
static char s_szDllPath[MAX_PATH + 1];
static BOOL s_bLog = FALSE;
static LONG s_nTlsIndent = -1;
static LONG s_nTlsThread = -1;
static LONG s_nThreadCnt = 0;

extern "C" {
	//  Trampolines for SYELOG library.
	//
	HANDLE(WINAPI *Real_CreateFileW)(LPCWSTR a0, DWORD a1, DWORD a2,
		LPSECURITY_ATTRIBUTES a3, DWORD a4, DWORD a5,
		HANDLE a6)
		= CreateFileW;
	BOOL(WINAPI *Real_WriteFile)(HANDLE hFile,
		LPCVOID lpBuffer,
		DWORD nNumberOfBytesToWrite,
		LPDWORD lpNumberOfBytesWritten,
		LPOVERLAPPED lpOverlapped)
		= WriteFile;
	BOOL(WINAPI *Real_FlushFileBuffers)(HANDLE hFile)
		= FlushFileBuffers;
	BOOL(WINAPI *Real_CloseHandle)(HANDLE hObject)
		= CloseHandle;
	BOOL(WINAPI *Real_WaitNamedPipeW)(LPCWSTR lpNamedPipeName, DWORD nTimeOut)
		= WaitNamedPipeW;
	BOOL(WINAPI *Real_SetNamedPipeHandleState)(HANDLE hNamedPipe,
		LPDWORD lpMode,
		LPDWORD lpMaxCollectionCount,
		LPDWORD lpCollectDataTimeout)
		= SetNamedPipeHandleState;
	DWORD(WINAPI *Real_GetCurrentProcessId)(VOID)
		= GetCurrentProcessId;
	VOID(WINAPI *Real_GetSystemTimeAsFileTime)(LPFILETIME lpSystemTimeAsFileTime)
		= GetSystemTimeAsFileTime;

	VOID(WINAPI * Real_InitializeCriticalSection)(LPCRITICAL_SECTION lpSection)
		= InitializeCriticalSection;
	VOID(WINAPI * Real_EnterCriticalSection)(LPCRITICAL_SECTION lpSection)
		= EnterCriticalSection;
	VOID(WINAPI * Real_LeaveCriticalSection)(LPCRITICAL_SECTION lpSection)
		= LeaveCriticalSection;
}

FARPROC(__stdcall *Real_GetProcAddress)(_In_ HMODULE hModule, _In_ LPCSTR lpProcName)
	= GetProcAddress;
SHORT(__stdcall *Real_GetAsyncKeyState)(_In_ int vKey)
	= GetAsyncKeyState;
void(__stdcall *Real_SysFreeString)(__in_opt BSTR bstrString)
	= SysFreeString;
BSTR(__stdcall *Real_SysAllocString)(__in_z_opt const OLECHAR * psz)
	= SysAllocString;
INT(__stdcall *Real_SysReAllocString)(__deref_inout_ecount_z(stringLength(psz) + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz)
	= SysReAllocString;
BSTR(__stdcall *Real_SysAllocStringLen)(__in_ecount_opt(ui) const OLECHAR * strIn, UINT ui)
	= SysAllocStringLen;
INT(__stdcall *Real_SysReAllocStringLen)(__deref_inout_ecount_z(len + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz, __in unsigned int len)
	= SysReAllocStringLen;
int(__stdcall *Real_GetWindowTextA)(__in HWND hWnd, __out_ecount(nMaxCount) LPSTR lpString, __in int nMaxCount)
	= GetWindowTextA;
BOOL(__stdcall *Real_SetWindowTextA)(_In_ HWND hWnd, _In_opt_ LPCSTR lpString)
	= SetWindowTextA;
BOOL(__stdcall *Real_SetWindowTextW)(_In_ HWND hWnd, _In_opt_ LPCWSTR lpString)
	= SetWindowTextW;
LRESULT(__stdcall *Real_DefWindowProcA)(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
	= DefWindowProcA;
LRESULT(__stdcall *Real_DefWindowProcW)(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
	= DefWindowProcW;
BOOL(__stdcall *Real_GetTextMetricsW)(__in HDC hdc, __out LPTEXTMETRICW lptm)
	= GetTextMetricsW;
BOOL(__stdcall *Real_ExtTextOutW)(__in HDC hdc, __in int x, __in int y, __in UINT options, __in_opt CONST RECT * lprect, __in_ecount_opt(c) LPCWSTR lpString, __in UINT c, __in_ecount_opt(c) CONST INT * lpDx)
	= ExtTextOutW;
HGDIOBJ (__stdcall *Real_SelectObject)(__in HDC hdc, __in HGDIOBJ h)
	= SelectObject;
BOOL(__stdcall *Real_SetWindowPos)(__in HWND hWnd, __in_opt HWND hWndInsertAfter, __in int X, __in int Y, __in int cx, __in int cy, __in UINT uFlags)
	= SetWindowPos;
LPVOID(__stdcall *Real_HeapAlloc)(_In_ HANDLE hHeap, _In_ DWORD dwFlags, _In_ SIZE_T dwBytes)
	= HeapAlloc;
BOOL(__stdcall *Real_PostMessageA)(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
	= PostMessageA;
BOOL (__stdcall *Real_PostMessageW)(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
	= PostMessageW;
LRESULT (__stdcall *Real_SendMessageA)(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam)
	= SendMessageA;
LRESULT (__stdcall *Real_SendMessageW)(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam)
	= SendMessageW;

FARPROC __stdcall Mine_GetProcAddress(HMODULE hModule, LPCSTR lpProcName);
SHORT __stdcall Mine_GetAsyncKeyState(_In_ int vKey);
void __stdcall Mine_SysFreeString(__in_opt BSTR bstrString);
BSTR __stdcall Mine_SysAllocString(__in_z_opt const OLECHAR * psz);
INT  __stdcall Mine_SysReAllocString(__deref_inout_ecount_z(stringLength(psz) + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz);
BSTR __stdcall Mine_SysAllocStringLen(__in_ecount_opt(ui) const OLECHAR * strIn, UINT ui);
INT __stdcall  Mine_SysReAllocStringLen(__deref_inout_ecount_z(len + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz, __in unsigned int len);
int __stdcall Mine_GetWindowTextA(__in HWND hWnd, __out_ecount(nMaxCount) LPSTR lpString, __in int nMaxCount);
BOOL __stdcall Mine_SetWindowTextA(_In_ HWND hWnd, _In_opt_ LPCSTR lpString);
BOOL __stdcall Mine_SetWindowTextW(_In_ HWND hWnd, _In_opt_ LPCWSTR lpString);
LRESULT __stdcall Mine_DefWindowProcA(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam);
LRESULT __stdcall Mine_DefWindowProcW(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam);
BOOL __stdcall Mine_GetTextMetricsW(__in HDC hdc, __out LPTEXTMETRICW lptm);
BOOL __stdcall Mine_ExtTextOutW(__in HDC hdc, __in int x, __in int y, __in UINT options, __in_opt CONST RECT * lprect, __in_ecount_opt(c) LPCWSTR lpString, __in UINT c, __in_ecount_opt(c) CONST INT * lpDx);
HGDIOBJ __stdcall Mine_SelectObject(__in HDC hdc, __in HGDIOBJ h);
BOOL __stdcall Mine_SetWindowPos(__in HWND hWnd, __in_opt HWND hWndInsertAfter, __in int X, __in int Y, __in int cx, __in int cy, __in UINT uFlags);
LPVOID __stdcall Mine_HeapAlloc(_In_ HANDLE hHeap, _In_ DWORD dwFlags, _In_ SIZE_T dwBytes);
BOOL __stdcall Mine_PostMessageA(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam);
BOOL __stdcall Mine_PostMessageW(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam);
LRESULT __stdcall Mine_SendMessageA(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam);
LRESULT __stdcall Mine_SendMessageW(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam);

ProcMdCallBack MdCallBack = NULL;
ProcMdCallBack12 MdCallBack12 = NULL;
Proc_LPenHelper _LPenHelper = NULL;

ProcMdCallBack Real_MdCallBack = NULL;
ProcMdCallBack12 Real_MdCallBack12 = NULL;
Proc_LPenHelper Real__LPenHelper = NULL;

#define Trace printf
BOOL WINAPI DllMain(
	HANDLE hinst,
	ULONG dwReason,
	LPVOID lpReserved)
{
	switch (dwReason) {
	case DLL_PROCESS_ATTACH:
	{
		TCHAR szMyPath[MAX_PATH];
		TCHAR szExePath[MAX_PATH];
		GetModuleFileName(NULL, szMyPath, MAX_PATH);
		GetModuleFileName(NULL, szExePath, MAX_PATH);
		TCHAR *p1 = wcsrchr(szMyPath, '\\');
		TCHAR *p2 = wcsrchr(szExePath, '\\');
		if (p2 && _wcsicmp(p2 + 1, L"ollyice.exe"))
		{
			s_hInst = (HMODULE)hinst;
			DetourRestoreAfterWith();
			ProcessAttach((HMODULE)hinst);
// 			LogHelper::Instance().OpenLogFile();
			s_hThis = ::LoadLibraryA("XllHook.xll");
		}
		else
		{
			return FALSE;
		}
	}
		break;
	case DLL_PROCESS_DETACH:
		LogHelper::Instance().CloseLogFile();
		ProcessDetach((HMODULE)hinst);
		break;
	case DLL_THREAD_ATTACH:
		ThreadAttach((HMODULE)hinst);
		break;
	case DLL_THREAD_DETACH:
		ThreadDetach((HMODULE)hinst);
		break;
	}

	return TRUE;
}

void __stdcall NullExport(){}

BOOL ThreadAttach(HMODULE hDll)
{
	(void)hDll;

	if (s_nTlsIndent >= 0) {
		TlsSetValue(s_nTlsIndent, (PVOID)0);
	}
	if (s_nTlsThread >= 0) {
		LONG nThread = InterlockedIncrement(&s_nThreadCnt);
		TlsSetValue(s_nTlsThread, (PVOID)(LONG_PTR)nThread);
	}
	return TRUE;
}

BOOL ThreadDetach(HMODULE hDll)
{
	(void)hDll;

	if (s_nTlsIndent >= 0) {
		TlsSetValue(s_nTlsIndent, (PVOID)0);
	}
	if (s_nTlsThread >= 0) {
		TlsSetValue(s_nTlsThread, (PVOID)0);
	}
	return TRUE;
}

PIMAGE_NT_HEADERS NtHeadersForInstance(HINSTANCE hInst)
{
	PIMAGE_DOS_HEADER pDosHeader = (PIMAGE_DOS_HEADER)hInst;
	__try {
		if (pDosHeader->e_magic != IMAGE_DOS_SIGNATURE) {
			SetLastError(ERROR_BAD_EXE_FORMAT);
			return NULL;
		}

		PIMAGE_NT_HEADERS pNtHeader = (PIMAGE_NT_HEADERS)((PBYTE)pDosHeader +
			pDosHeader->e_lfanew);
		if (pNtHeader->Signature != IMAGE_NT_SIGNATURE) {
			SetLastError(ERROR_INVALID_EXE_SIGNATURE);
			return NULL;
		}
		if (pNtHeader->FileHeader.SizeOfOptionalHeader == 0) {
			SetLastError(ERROR_EXE_MARKED_INVALID);
			return NULL;
		}
		return pNtHeader;
	}
	__except (EXCEPTION_EXECUTE_HANDLER) {
	}
	SetLastError(ERROR_EXE_MARKED_INVALID);

	return NULL;
}

BOOL InstanceEnumerate(HINSTANCE hInst)
{
	WCHAR wzDllName[MAX_PATH];

	PIMAGE_NT_HEADERS pinh = NtHeadersForInstance(hInst);
	if (pinh && GetModuleFileNameW(hInst, wzDllName, ARRAYSIZE(wzDllName))) {
		Syelog(SYELOG_SEVERITY_INFORMATION, "### %p: %ls\n", hInst, wzDllName);
		return TRUE;
	}
	return FALSE;
}

BOOL ProcessEnumerate()
{
	Syelog(SYELOG_SEVERITY_INFORMATION,
		"######################################################### Binaries\n");

	PBYTE pbNext;
	for (PBYTE pbRegion = (PBYTE)0x10000;; pbRegion = pbNext) {
		MEMORY_BASIC_INFORMATION mbi;
		ZeroMemory(&mbi, sizeof(mbi));

		if (VirtualQuery((PVOID)pbRegion, &mbi, sizeof(mbi)) <= 0) {
			break;
		}
		pbNext = (PBYTE)mbi.BaseAddress + mbi.RegionSize;

		// Skip free regions, reserver regions, and guard pages.
		//
		if (mbi.State == MEM_FREE || mbi.State == MEM_RESERVE) {
			continue;
		}
		if (mbi.Protect & PAGE_GUARD || mbi.Protect & PAGE_NOCACHE) {
			continue;
		}
		if (mbi.Protect == PAGE_NOACCESS) {
			continue;
		}

		// Skip over regions from the same allocation...
		{
			MEMORY_BASIC_INFORMATION mbiStep;

			while (VirtualQuery((PVOID)pbNext, &mbiStep, sizeof(mbiStep)) > 0) {
				if ((PBYTE)mbiStep.AllocationBase != pbRegion) {
					break;
				}
				pbNext = (PBYTE)mbiStep.BaseAddress + mbiStep.RegionSize;
				mbi.Protect |= mbiStep.Protect;
			}
		}

		WCHAR wzDllName[MAX_PATH];
		PIMAGE_NT_HEADERS pinh = NtHeadersForInstance((HINSTANCE)pbRegion);

		if (pinh &&
			GetModuleFileNameW((HINSTANCE)pbRegion, wzDllName, ARRAYSIZE(wzDllName))) {

			Syelog(SYELOG_SEVERITY_INFORMATION,
				"### %p..%p: %ls\n", pbRegion, pbNext, wzDllName);
		}
		else {
			Syelog(SYELOG_SEVERITY_INFORMATION,
				"### %p..%p: State=%04x, Protect=%08x\n",
				pbRegion, pbNext, mbi.State, mbi.Protect);
		}
	}
	Syelog(SYELOG_SEVERITY_INFORMATION, "###\n");

	LPVOID lpvEnv = GetEnvironmentStrings();
	Syelog(SYELOG_SEVERITY_INFORMATION, "### Env= %08x [%08x %08x]\n",
		lpvEnv, ((PVOID*)lpvEnv)[0], ((PVOID*)lpvEnv)[1]);

	return TRUE;
}

BOOL ProcessAttach(HMODULE hDll)
{
	s_bLog = FALSE;
	s_nTlsIndent = TlsAlloc();
	s_nTlsThread = TlsAlloc();
	ThreadAttach(hDll);

	WCHAR wzExeName[MAX_PATH];

	s_hInst = hDll;
	s_hExcel = GetModuleHandle(NULL);
	GetModuleFileNameW(hDll, s_wzDllPath, ARRAYSIZE(s_wzDllPath));
	GetModuleFileNameW(NULL, wzExeName, ARRAYSIZE(wzExeName));
	sprintf_s(s_szDllPath, ARRAYSIZE(s_szDllPath), "%ls", s_wzDllPath);

	SyelogOpen("trcapi" DETOURS_STRINGIFY(DETOURS_BITS), SYELOG_FACILITY_APPLICATION);
	ProcessEnumerate();

	LONG error = AttachDetours();
	if (error != NO_ERROR) {
		Syelog(SYELOG_SEVERITY_FATAL, "### Error attaching detours: %d\n", error);
	}

	return TRUE;
}

BOOL ProcessDetach(HMODULE hDll)
{
	ThreadDetach(hDll);
	s_bLog = FALSE;

	LONG error = DetachDetours();
	if (error != NO_ERROR) {
		Syelog(SYELOG_SEVERITY_FATAL, "### Error detaching detours: %d\n", error);
	}

	Syelog(SYELOG_SEVERITY_NOTICE, "### Closing.\n");
	SyelogClose(FALSE);

	if (s_nTlsIndent >= 0) {
		TlsFree(s_nTlsIndent);
	}
	if (s_nTlsThread >= 0) {
		TlsFree(s_nTlsThread);
	}
	return TRUE;
}

static PCHAR DetRealName(PCHAR psz)
{
	PCHAR pszBeg = psz;
	// Move to end of name.
	while (*psz) {
		psz++;
	}
	// Move back through A-Za-z0-9 names.
	while (psz > pszBeg &&
		((psz[-1] >= 'A' && psz[-1] <= 'Z') ||
		(psz[-1] >= 'a' && psz[-1] <= 'z') ||
		(psz[-1] >= '0' && psz[-1] <= '9'))) {
		psz--;
	}
	return psz;
}

static VOID Dump(PBYTE pbBytes, LONG nBytes, PBYTE pbTarget)
{
	CHAR szBuffer[256];
	PCHAR pszBuffer = szBuffer;

	for (LONG n = 0; n < nBytes; n += 12) {
#ifdef _CRT_INSECURE_DEPRECATE
		pszBuffer += sprintf_s(pszBuffer, sizeof(szBuffer), "  %p: ", pbBytes + n);
#else
		pszBuffer += sprintf(pszBuffer, "  %p: ", pbBytes + n);
#endif
		for (LONG m = n; m < n + 12; m++) {
			if (m >= nBytes) {
#ifdef _CRT_INSECURE_DEPRECATE
				pszBuffer += sprintf_s(pszBuffer, sizeof(szBuffer), "   ");
#else
				pszBuffer += sprintf(pszBuffer, "   ");
#endif
			}
			else {
#ifdef _CRT_INSECURE_DEPRECATE
				pszBuffer += sprintf_s(pszBuffer, sizeof(szBuffer), "%02x ", pbBytes[m]);
#else
				pszBuffer += sprintf(pszBuffer, "%02x ", pbBytes[m]);
#endif
			}
		}
		if (n == 0) {
#ifdef _CRT_INSECURE_DEPRECATE
			pszBuffer += sprintf_s(pszBuffer, sizeof(szBuffer), "[%p]", pbTarget);
#else
			pszBuffer += sprintf(pszBuffer, "[%p]", pbTarget);
#endif
		}
#ifdef _CRT_INSECURE_DEPRECATE
		pszBuffer += sprintf_s(pszBuffer, sizeof(szBuffer), "\n");
#else
		pszBuffer += sprintf(pszBuffer, "\n");
#endif
	}

	Syelog(SYELOG_SEVERITY_INFORMATION, "%s", szBuffer);
}

static VOID Decode(PBYTE pbCode, LONG nInst)
{
	PBYTE pbSrc = pbCode;
	PBYTE pbEnd;
	PBYTE pbTarget;
	for (LONG n = 0; n < nInst; n++) {
		pbTarget = NULL;
		pbEnd = (PBYTE)DetourCopyInstruction(NULL, NULL, (PVOID)pbSrc, (PVOID*)&pbTarget, NULL);
		Dump(pbSrc, (int)(pbEnd - pbSrc), pbTarget);
		pbSrc = pbEnd;

		if (pbTarget != NULL) {
			break;
		}
	}
}

VOID DetAttach(PVOID *ppvReal, PVOID pvMine, PCHAR psz)
{
	PVOID pvReal = NULL;
	if (ppvReal == NULL) {
		ppvReal = &pvReal;
	}

	LONG l = DetourAttach(ppvReal, pvMine);
	if (l != 0) {
		Syelog(SYELOG_SEVERITY_NOTICE,
			"Attach failed: `%s': error %d\n", DetRealName(psz), l);

		Decode((PBYTE)*ppvReal, 3);
	}
}

VOID DetDetach(PVOID *ppvReal, PVOID pvMine, PCHAR psz)
{
	LONG l = DetourDetach(ppvReal, pvMine);
	if (l != 0) {
#if 0
		Syelog(SYELOG_SEVERITY_NOTICE,
			"Detach failed: `%s': error %d\n", DetRealName(psz), l);
#else
		(void)psz;
#endif
	}
}

#define ATTACH(x)       DetAttach(&(PVOID&)Real_##x,Mine_##x,#x)
#define DETACH(x)       DetDetach(&(PVOID&)Real_##x,Mine_##x,#x)

LONG AttachDetours(VOID)
{
	Real_MdCallBack = MdCallBack = (ProcMdCallBack)DetourFindFunction("EXCEL.EXE", Excel_MdCallBack);
	Real_MdCallBack12 = MdCallBack12 = (ProcMdCallBack12)DetourFindFunction("EXCEL.EXE", Excel_MdCallBack12);
	Real__LPenHelper = _LPenHelper = (Proc_LPenHelper)DetourFindFunction("EXCEL.EXE", Excel_LPenHelper);
	if (!MdCallBack && !MdCallBack12 && !_LPenHelper)
	{
		Real_MdCallBack = MdCallBack = (ProcMdCallBack)DetourFindFunction("et.EXE", Excel_MdCallBack);
		Real_MdCallBack12 = MdCallBack12 = (ProcMdCallBack12)DetourFindFunction("et.EXE", Excel_MdCallBack12);
		Real__LPenHelper = _LPenHelper = (Proc_LPenHelper)DetourFindFunction("et.EXE", Excel_LPenHelper);
	}
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());

	// For this many APIs, we'll ignore one or two can't be detoured.
	DetourSetIgnoreTooSmall(TRUE);

#if SET_Hook_XLL
	if (MdCallBack)
		ATTACH(MdCallBack);
	if (MdCallBack12)
		ATTACH(MdCallBack12);
	if (_LPenHelper)
		ATTACH(_LPenHelper);
#endif
#if SET_Hook_Other
// 	ATTACH(GetProcAddress);
// 	ATTACH(GetAsyncKeyState);
// 	ATTACH(DefWindowProcA);
// 	ATTACH(DefWindowProcW);
// 	ATTACH(GetWindowTextA);
// 	ATTACH(SetWindowTextA);
// 	ATTACH(SetWindowTextW);
//  ATTACH(SysFreeString);
// 	ATTACH(SysAllocString);
// 	ATTACH(SysReAllocString);
// 	ATTACH(SysAllocStringLen);
// 	ATTACH(SysReAllocStringLen);
// 	ATTACH(GetTextMetricsW);
//	ATTACH(ExtTextOutW);
// 	ATTACH(HeapAlloc);
// 	ATTACH(SelectObject);
// 	ATTACH(SetWindowPos);
//  ATTACH(PostMessageA);
// 	ATTACH(PostMessageW);
// 	ATTACH(SendMessageA);
// 	ATTACH(SendMessageW);
#endif
	if (DetourTransactionCommit() != NO_ERROR) {
		OutputDebugStringA("AttachDetours failed on DetourTransactionCommit\n");

		PVOID *ppbFailedPointer = NULL;
		LONG error = DetourTransactionCommitEx(&ppbFailedPointer);

// 		Trace("DetourTransactionCommitEx Error: %d (%p/%p)", error, ppbFailedPointer, *ppbFailedPointer);
		return error;
	}
	else{
		OutputDebugStringA("AttachDetours OK\n");
	}

	return 0;
}

LONG DetachDetours(VOID)
{
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());

	// For this many APIs, we'll ignore one or two can't be detoured.
	DetourSetIgnoreTooSmall(TRUE);

#if SET_Hook_XLL
	if (MdCallBack)
		DETACH(MdCallBack);
	if (MdCallBack12)
		DETACH(MdCallBack12);
	if (_LPenHelper)
		DETACH(_LPenHelper);
#endif
#if SET_Hook_Other
// 	DETACH(GetProcAddress);
// 	DETACH(GetAsyncKeyState);
// 	DETACH(DefWindowProcA);
// 	DETACH(DefWindowProcW);
// 	DETACH(GetWindowTextA);
// 	DETACH(SetWindowTextA);
// 	DETACH(SetWindowTextW);
//  DETACH(SysFreeString);
// 	DETACH(SysAllocString);
// 	DETACH(SysReAllocString);
// 	DETACH(SysAllocStringLen);
// 	DETACH(SysReAllocStringLen);
//  DETACH(GetTextMetricsW);
//	DETACH(ExtTextOutW);
// 	DETACH(HeapAlloc);
// 	DETACH(SelectObject);
// 	DETACH(SetWindowPos);
// 	DETACH(PostMessageA);
// 	DETACH(PostMessageW);
// 	DETACH(SendMessageA);
// 	DETACH(SendMessageW);
#endif
	if (DetourTransactionCommit() != 0) {
		PVOID *ppbFailedPointer = NULL;
		LONG error = DetourTransactionCommitEx(&ppbFailedPointer);

		printf("traceapi.dll: Detach transaction failed to commit. Error %d (%p/%p)",
			error, ppbFailedPointer, *ppbFailedPointer);
		return error;
	}
	return 0;
}

void AttachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz)
{
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());

	// For this many APIs, we'll ignore one or two can't be detoured.
	DetourSetIgnoreTooSmall(TRUE);

	DetAttach(ppvReal, pvMine, psz);

	if (DetourTransactionCommit() != NO_ERROR) {
		OutputDebugStringA("AttachDetours failed on DetourTransactionCommit\n");

		PVOID *ppbFailedPointer = NULL;
		LONG error = DetourTransactionCommitEx(&ppbFailedPointer);

		// 		Trace("DetourTransactionCommitEx Error: %d (%p/%p)", error, ppbFailedPointer, *ppbFailedPointer);
		return;
	}
	else{
		OutputDebugStringA("AttachDetours OK\n");
	}

}

void DetachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz)
{
	DetourTransactionBegin();
	DetourUpdateThread(GetCurrentThread());

	// For this many APIs, we'll ignore one or two can't be detoured.
	DetourSetIgnoreTooSmall(TRUE);

	DetDetach(ppvReal, pvMine, psz);

	if (DetourTransactionCommit() != 0) {
		PVOID *ppbFailedPointer = NULL;
		LONG error = DetourTransactionCommitEx(&ppbFailedPointer);

		printf("traceapi.dll: Detach transaction failed to commit. Error %d (%p/%p)",
			error, ppbFailedPointer, *ppbFailedPointer);
	}
}

#pragma region WindowsApi

FARPROC __stdcall Mine_GetProcAddress(HMODULE hModule, LPCSTR lpProcName)
{
	FARPROC rv = 0;
	__try {
		rv = Real_GetProcAddress(hModule, lpProcName);
		if (!IsBadReadPtr(lpProcName, sizeof(void*)))
		{
		}
	}
	__finally {
		// 		_PrintExit("GetProcAddress(,) -> %p\n", rv);
	};
	return rv;
}

SHORT __stdcall Mine_GetAsyncKeyState(_In_ int vKey)
{
	SHORT res = Real_GetAsyncKeyState(vKey);
	if (vKey == VK_CANCEL)
	{
		int i = 0;
		++i;
	}
	return res;
}

std::map<const BSTR, int> g_addrMap;
bool sysAllocLock = false;

bool IsTarget(BSTR bstrString)
{
	return true;
	if (g_addrMap.find(bstrString) != g_addrMap.end())
		return true;

	const WCHAR* target1 = L"DESSeal.DESSealObj.1";
	if (!IsBadReadPtr(bstrString, sizeof(void*)))
	{
		return (0 == _wcsicmp(target1, bstrString));
	}
	return false;
}

void __stdcall Mine_SysFreeString(__in_opt BSTR bstrString)
{
	if (IsTarget(bstrString))
	{
// 		assert(g_addrMap[bstrString] > 0);
		--g_addrMap[bstrString];
		if (g_addrMap[bstrString] < 0)
		{
			DebugBreak();
			OutputDebugStringW(bstrString);
			OutputDebugStringA("\n");
		}
// 		assert(0 == g_addrMap[bstrString]);
// 		if (g_addrMap[bstrString] < 0)
// 			g_addrMap[bstrString] = 0;
	}
	Real_SysFreeString(bstrString);
}

BSTR __stdcall Mine_SysAllocString(__in_z_opt const OLECHAR * psz)
{
	sysAllocLock = true;
	BSTR bstrString = Real_SysAllocString(psz);
	sysAllocLock = false;

	if (IsTarget(bstrString))
	{
// 		assert(g_addrMap[bstrString] >= 0);
		++g_addrMap[bstrString];
	}
	return bstrString;
}

INT __stdcall Mine_SysReAllocString(__deref_inout_ecount_z(stringLength(psz) + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz)
{
	if (IsTarget(*pbstr))
	{
// 		assert(g_addrMap[*pbstr] > 0);
		--g_addrMap[*pbstr];
// 		if (g_addrMap[*pbstr] < 0)
// 			g_addrMap[*pbstr] = 0;
	}
	sysAllocLock = true;
	INT res = Real_SysReAllocString(pbstr, psz);
	sysAllocLock = false;

	if (IsTarget(*pbstr))
	{
// 		assert(g_addrMap[*pbstr] >= 0);
		++g_addrMap[*pbstr];
	}
	return res;
}

BSTR __stdcall Mine_SysAllocStringLen(__in_ecount_opt(ui) const OLECHAR * strIn, UINT ui)
{
	BSTR bstrString = Real_SysAllocStringLen(strIn, ui);

	if (!sysAllocLock && IsTarget(bstrString))
	{
// 		assert(g_addrMap[bstrString] >= 0);
		++g_addrMap[bstrString];
	}
	return bstrString;
}

INT __stdcall Mine_SysReAllocStringLen(__deref_inout_ecount_z(len + 1) BSTR* pbstr, __in_z_opt const OLECHAR* psz, __in unsigned int len)
{
	if (!sysAllocLock && g_addrMap.find(*pbstr) != g_addrMap.end())
	{
// 		assert(g_addrMap[*pbstr] > 0);
		--g_addrMap[*pbstr];
// 		if (g_addrMap[*pbstr] < 0)
// 			g_addrMap[*pbstr] = 0;
	}

	INT res = Real_SysReAllocStringLen(pbstr, psz, len);

	if (!sysAllocLock && IsTarget(*pbstr))
	{
// 		assert(g_addrMap[*pbstr] >= 0);
		++g_addrMap[*pbstr];
	}
	return res;
}

int __stdcall Mine_GetWindowTextA(__in HWND hWnd, __out_ecount(nMaxCount) LPSTR lpString, __in int nMaxCount)
{
	int len = Real_GetWindowTextA(hWnd, lpString, nMaxCount);
	//int len = DefWindowProcA(hWnd, WM_GETTEXT, nMaxCount, (LPARAM)lpString);
	if (len > 0xFF)
	{
		len = 0;
	}
	else
	{
		int tmp = strnlen(lpString, nMaxCount);
		if (tmp < len)
			len = tmp;
	}
	return len;
}

BOOL __stdcall Mine_SetWindowTextA(_In_ HWND hWnd, _In_opt_ LPCSTR lpString)
{
	return Real_SetWindowTextA(hWnd, lpString);
}

BOOL __stdcall Mine_SetWindowTextW(_In_ HWND hWnd, _In_opt_ LPCWSTR lpString)
{
	return Real_SetWindowTextW(hWnd, lpString);
}

LRESULT __stdcall Mine_DefWindowProcA(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
{
	if (Msg == WM_GETTEXT)
	{
		return Real_DefWindowProcA(hWnd, Msg, wParam, lParam);
	}
	return Real_DefWindowProcA(hWnd, Msg, wParam, lParam);
}

LRESULT __stdcall Mine_DefWindowProcW(_In_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
{
	if (Msg == WM_GETTEXT)
	{
		return Real_DefWindowProcW(hWnd, Msg, wParam, lParam);
	}
	return Real_DefWindowProcW(hWnd, Msg, wParam, lParam);
}

BOOL __stdcall Mine_GetTextMetricsW(__in HDC hdc, __out LPTEXTMETRICW lptm)
{
	BOOL res = Real_GetTextMetricsW(hdc, lptm);
	return res;
	static int ascOff = 0;
	static int desOff = 0;
	static int extOff = 0;
	static int intOff = 0;
	if (res)
	{
		int nDPI = GetDeviceCaps(hdc, LOGPIXELSY);
		double fDesktopUintPerPels = 15;	// 96 dpi
		double fDeviceUintPerPels = static_cast<double>(1440 * 1.0 / nDPI);
		double nRate = fDesktopUintPerPels / fDeviceUintPerPels;

		static WCHAR fontName[255];
		GetTextFaceW(hdc, 255, fontName);

		ASSERT(lptm->tmAscent + lptm->tmDescent == lptm->tmHeight);
		lptm->tmAscent += ascOff;
		lptm->tmDescent += desOff;
		lptm->tmHeight = lptm->tmAscent + lptm->tmDescent;
		lptm->tmExternalLeading += extOff;
		lptm->tmInternalLeading += intOff;

		if (lptm->tmAscent == 75)
		{
			return res;
		}
	}
	return res;
}

BOOL __stdcall Mine_ExtTextOutW(
	__in HDC hdc, __in int x, __in int y,
	__in UINT options, __in_opt CONST RECT * lprect,
	__in_ecount_opt(c) LPCWSTR lpString, __in UINT c, __in_ecount_opt(c) CONST INT * lpDx)
{
// 	const UINT firstLine = 0x91cd3000;
// 	const UINT secondLine = 0x66f87d04;
	const UINT firstLine = *(UINT*)&L"LL";
	const UINT secondLine = *(UINT*)&L"AA";

	if (!IsBadReadPtr(lpString, sizeof(LPCWSTR)) &&
		!IsBadReadPtr(lprect, sizeof(RECT *)))
	{

		const WCHAR* lpTarget = L"1999.07";
		if (0 == memcmp(lpString, lpTarget, 7))
		{
			TEXTMETRICW tm = { 0 };
			GetTextMetricsW(hdc, &tm);
			OUTLINETEXTMETRIC olMetric = { 0 };
			GetOutlineTextMetricsW(hdc, sizeof(olMetric), &olMetric);
			static WCHAR nameBuff[255];
			GetTextFaceW(hdc, 255, nameBuff);

			static WCHAR buff[255];
			swprintf_s(buff, L"Height(%d) MaxWidth(%d) Name(%s)\n\0",
				tm.tmHeight, olMetric.otmSize, nameBuff);
			OutputDebugStringW(buff);
		}
	}
	return Real_ExtTextOutW(hdc, x, y, options, lprect, lpString, c, lpDx);
}

HGDIOBJ __stdcall Mine_SelectObject(__in HDC hdc, __in HGDIOBJ h)
{
#define PrintCase(name) case name: OutputDebugStringA( #name "\n"); break;
	DWORD objType = GetObjectType(h);
	switch (objType)
	{
		PrintCase(OBJ_BITMAP);
		PrintCase(OBJ_BRUSH);
		PrintCase(OBJ_COLORSPACE);
		PrintCase(OBJ_DC);
		PrintCase(OBJ_ENHMETADC);
		PrintCase(OBJ_ENHMETAFILE);
		PrintCase(OBJ_EXTPEN);
		PrintCase(OBJ_FONT);
		PrintCase(OBJ_MEMDC);
		PrintCase(OBJ_METAFILE);
		PrintCase(OBJ_METADC);
		PrintCase(OBJ_PAL);
		PrintCase(OBJ_PEN);
		PrintCase(OBJ_REGION);
	default:
		break;
	}
	return Real_SelectObject(hdc, h);
}

BOOL __stdcall Mine_SetWindowPos(__in HWND hWnd, __in_opt HWND hWndInsertAfter, __in int X, __in int Y, __in int cx, __in int cy, __in UINT uFlags)
{
	return Real_SetWindowPos(hWnd, hWndInsertAfter, X, Y, cx, cy, uFlags);
}

LPVOID __stdcall Mine_HeapAlloc(_In_ HANDLE hHeap, _In_ DWORD dwFlags, _In_ SIZE_T dwBytes)
{
	static LPVOID lpBegin = 0;
	static LPVOID lpEnd = 0;
	LPVOID lpRes = Real_HeapAlloc(hHeap, dwFlags, dwBytes);
	if (lpBegin <= lpRes && lpRes <= lpEnd)
	{
		static UINT i = 0;
		++i;
	}
	return lpRes;
}

BOOL __stdcall Mine_PostMessageA(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
{
	if (Msg == 2224 || Msg == 2225)
	{
		OutputDebugStringA("hint");
	}
	return Real_PostMessageA(hWnd, Msg, wParam, lParam);
}
BOOL __stdcall Mine_PostMessageW(_In_opt_ HWND hWnd, _In_ UINT Msg, _In_ WPARAM wParam, _In_ LPARAM lParam)
{
	if (Msg == 2224 || Msg == 2225)
	{
		OutputDebugStringA("hint");
	}
	return Real_PostMessageW(hWnd, Msg, wParam, lParam);
}

LRESULT __stdcall Mine_SendMessageA(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam)
{
	if (Msg == 2224 || Msg == 2225)
	{
		OutputDebugStringA("hint");
	}
	return Real_SendMessageA(hWnd, Msg, wParam, lParam);
}

LRESULT __stdcall Mine_SendMessageW(_In_ HWND hWnd, _In_ UINT Msg, _Pre_maybenull_ _Post_valid_ WPARAM wParam, _Pre_maybenull_ _Post_valid_ LPARAM lParam)
{
	if (Msg == 2224 || Msg == 2225)
	{
		OutputDebugStringA("hint");
	}
	return Real_SendMessageW(hWnd, Msg, wParam, lParam);
}

#pragma endregion
