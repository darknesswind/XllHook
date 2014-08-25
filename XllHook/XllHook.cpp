#include <ctype.h>
#include <windows.h>
#include "detver.h"
#include "detours.h"
#include "syelog.h"
#include <cstdio>
#include <cassert>
#include "ExcelProcxy.h"
#include "loghelper.h"
#include "xlcall.h"
#include "XllHook.h"

BOOL ProcessAttach(HMODULE hDll);
BOOL ProcessDetach(HMODULE hDll);
LONG AttachDetours(VOID);
LONG DetachDetours(VOID);
BOOL ThreadAttach(HMODULE hDll);
BOOL ThreadDetach(HMODULE hDll);

static HMODULE s_hExcel = NULL;
static HMODULE s_hInst = NULL;
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

FARPROC __stdcall Mine_GetProcAddress(HMODULE hModule, LPCSTR lpProcName);
SHORT __stdcall Mine_GetAsyncKeyState(_In_ int vKey);

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

// 	ATTACH(GetProcAddress);
	if (MdCallBack)
		ATTACH(MdCallBack);
	if (MdCallBack12)
		ATTACH(MdCallBack12);
	if (_LPenHelper)
		ATTACH(_LPenHelper);
// 	ATTACH(GetAsyncKeyState);

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

// 	DETACH(GetProcAddress);
	if (MdCallBack)
		DETACH(MdCallBack);
	if (MdCallBack12)
		DETACH(MdCallBack12);
	if (_LPenHelper)
		DETACH(_LPenHelper);
// 	DETACH(GetAsyncKeyState);

	if (DetourTransactionCommit() != 0) {
		PVOID *ppbFailedPointer = NULL;
		LONG error = DetourTransactionCommitEx(&ppbFailedPointer);

		printf("traceapi.dll: Detach transaction failed to commit. Error %d (%p/%p)",
			error, ppbFailedPointer, *ppbFailedPointer);
		return error;
	}
	return 0;
}

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
