#ifndef __LOG_HELPER_H__
#define __LOG_HELPER_H__

#include <vector>
#include <fstream>
#include <mutex>
#include <windows.h>
#include "xlcall.h"

#ifndef ASSERT
#define ASSERT(x) assert(x)
#endif
#define __X(x) L ## x
#define __Xc(x) L ## x

#ifndef STRINGIFY
#define STRINGIFY(x)   STRINGIFY_(x)
#define STRINGIFY_(x)  #x
#endif

#define Excel4MaxRow 65536
#define Excel4MaxCol 256
#define PascalStrMaxLen 0xFF
#define PascalWStrMaxLen 0xFFFF
#define FuncNumMask (0xFFF | xlCommand | xlSpecial)
#define FuncTypeMask (0xF000)
#define XLOPER_TYPEMASK (unsigned short)(0xFFF)

struct LogBuffer
{
	std::wstring sFuncAttr;
	std::wstring sFuncName;
	std::wstring sResult;
	std::wstring sResOperType;
	std::wstring sResOperValue;
	std::vector<std::wstring> argsOperType;
	std::vector<std::wstring> argsOperValue;

	LogBuffer()
	{
	}

	void clear()
	{
		sFuncAttr.clear();
		sFuncName.clear();
		sResult.clear();
		sResOperType.clear();
		sResOperValue.clear();
		argsOperType.clear();
		argsOperValue.clear();
	}
};

class LogHelper
{
public:
	LogHelper();

	static LogHelper& Instance() { return g_Instance; }

	void OpenLogFile();
	void CloseLogFile();

	void PauseLog() { m_bPause = true; }
	void ResumeLog() { m_bPause = false; }

	template <class LPOperType>
	void LogFunctionBegin(int xlfn, int coper, LPOperType *rgpxloper, LogBuffer& buffer);
	template <class LPOperType>
	void LogFunctionEnd(int result, LPOperType xloperRes, LogBuffer& buffer);

	void LogLPenHelperBegin(int wCode, void* lpv, LogBuffer& buffer);
	void LogLPenHelperEnd(int result, LogBuffer& buffer);

protected:
	void GetXlFunctionName(int xlfn, std::wstring& str);
	void GetXlFunctionTypeStr(int xlfn, std::wstring& str);
	void GetXlResultName(XLCALL_RESULT res, std::wstring& str);
	void GetXloperTypeName(int type, std::wstring& str);
	void GetXloperErrName(XLOPER_ERRTYPE type, std::wstring& str);
	void GetPascalString(LPCSTR, std::wstring& result);
	void GetPascalString(LPCWSTR, std::wstring& result);

	template <class LPOperType>
	void LogXloper(LPOperType lpOper, std::wstring& sType, std::wstring& sVal);

	template <class XLRefType>
	void LogSingleRef(const XLRefType& sref, std::wstringstream& stream);

	template <class XLMRefType>
	void LogReferences(IDSHEET iSheet, const XLMRefType* lpmref, std::wstringstream& stream);

	template <class LPOperType>
	void LogXloperFlow(LPOperType lpOper, std::wstringstream& stream);

	template <class LPOperType>
	void LogOperArray(RW row, COL col, LPOperType lpArray, std::wstringstream& stream);

	template <class LPOperType>
	void LogArrayToFile(RW row, COL col, LPOperType lpArray, UINT id);

	void PrintLogTitle();
	void PrintBuffer(LogBuffer& buffer);

private:
	UINT64 m_nLineCount;
	UINT m_nArrayCount;
	std::mutex	m_logFileMutex;
	std::mutex	m_arrayMutex;
	bool m_bPause;

	std::wstring m_sLogPath;
	std::wstring m_sLogFile;

	std::wofstream m_fileStream;

	static LogHelper g_Instance;
};

#include "loghelper.inl"

#endif