#ifndef __LOG_HELPER_H__
#define __LOG_HELPER_H__

#include <vector>
#include <fstream>
#include <mutex>
#include <windows.h>
#include "xlcall.h"
#include <map>

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

#define TableBegin __X("<table border=\"1\">")
#define TableEnd __X("</table>")
#define RowBegin __X("<tr>")
#define RowEnd __X("</tr>")
#define ColBegin __X("<td>")
#define ColEnd __X("</td>")

struct LogBuffer
{
	std::wstring sFuncAttr;
	std::wstring sFuncName;
	std::wstring sResult;
	std::wstring sResOperType;
	std::wstring sResOperValue;
	std::vector<std::wstring> argsOperType;
	std::vector<std::wstring> argsOperValue;

	int xlfn;
	bool bPrintEnter;
	LogBuffer() : xlfn(0), bPrintEnter(true)
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

enum XlCallArgType
{
	xlArgNone = 0,
	xlArgRetrun1 = 1,
	xlArgRetrun2 = 2,
	xlArgRetrun3 = 3,
	xlArgRetrun4 = 4,
	xlArgRetrun5 = 5,
	xlArgRetrun6 = 6,
	xlArgRetrun7 = 7,
	xlArgRetrun8 = 8,
	xlArgRetrun9 = 9,
	// <= excel 2003
	xlArgBool,		// A	bool
	xlArgBoolRef,	// L	bool*
	xlArgDouble,	// B	double
	xlArgDoubleRef,	// E	double*
	xlArgCStr,		// C	C����ַ�����\0��β��
	xlArgPascalStr,	// D	Pascal����ַ���(��һ���ַ�Ϊ�ַ�������)
	xlArgUShort,	// H	unsigned short
	xlArgShort,		// I	short
	xlArgShortRef,	// M	short*
	xlArgInt,		// J	int
	xlArgIntRef,	// N	int*
	xlArgFloatArr,	// K	�ṹ�� _FP
	xlArgArray,		// O	(ushort*, ushort*, double*)
	xlArgOper,		// P	xltypeNum, xltypeStr, xltypeBool, xltypeErr, xltypeMulti, xltypeMissing, xltypeNil
	xlArgXLOper,	// R	xltypeNum, xltypeStr, xltypeBool, xltypeErr, xltypeMulti, xltypeMissing, xltypeNil, xltypeRef, xltypeSRef
	// >= excel 2007
	xlArgCWStr,			// C%	Unicode C����ַ���
	xlArgPascalWStr,	// D%	Unicode Pascal����ַ���
	xlArgFloatArr12,// K%	�ṹ��_FP12
	xlArgArray12,	// O%	(int*, int*, double*)
	xlArgOper12,	// Q	xltypeNum, xltypeStr, xltypeBool, xltypeErr, xltypeMulti, xltypeMissing, xltypeNil
	xlArgXLOper12,	// U	xltypeNum, xltypeStr, xltypeBool, xltypeErr, xltypeMulti, xltypeMissing, xltypeNil, xltypeRef, xltypeSRef

	xlArgBitVolatile = 0x100, // !	recalculates every time the worksheet recalculates
	xlArgBitMacroFunc = 0x200,	// #	������R��U�����ͻ�Ĭ��Ϊ�ɱ��
	xlArgBitThreadSafe = 0x400, // $	��#������
	xlArgBitClusterSafe = 0x800, // &
	xlArgBitInPlaceModify = 0x1000,

	xlArgCStrInOut = xlArgCStr | xlArgBitInPlaceModify,			// F	ԭ�ر༭��C����ַ���
	xlArgCWStrInOut = xlArgCWStr | xlArgBitInPlaceModify,			// F%	ԭ�ر༭��Unicode C����ַ���
	xlArgPascalStrInOut = xlArgPascalStr | xlArgBitInPlaceModify,	// G	ԭ�ر༭��Pascal����ַ���
	xlArgPascalWStrInOut = xlArgPascalWStr | xlArgBitInPlaceModify,	// G%	ԭ�ر༭��Unicode Pascal����ַ���
};

class ShellCode
{
public:
	ShellCode(PVOID srcFunc = 0);
	PVOID address();
	void operator=(const ShellCode& other);

private:
	const static int m_size = 16;
	char m_code[m_size];
};

struct XllFuncInfo
{
	void* pEntryPoint;
	DWORD funcAttr;
	XlCallArgType retrunType;
	std::wstring funcName;

	std::vector<XlCallArgType> argTypes;
	XllFuncInfo()
		: pEntryPoint(NULL)
		, funcAttr(0)
		, retrunType(xlArgNone)
	{
	}
};

union XlFuncResult
{
	double dbl;
	DWORD dw;
};

typedef std::map<void*, XllFuncInfo> UDFMap;

class LogHelper
{
public:
	LogHelper();
	~LogHelper();
	static LogHelper& Instance() { return g_Instance; }

	void OpenLogFile();
	void CloseLogFile();

	void PauseLog() { m_bPause = true; }
	void ResumeLog() { m_bPause = false; }
	void ClearLog();
	void OpenFolder();

	template <class LPOperType>
	void LogFunctionBegin(int xlfn, int coper, LPOperType *rgpxloper);
	template <class LPOperType>
	void LogFunctionEnd(int result, LPOperType xloperRes);

	void LogLPenHelperBegin(int wCode, void* lpv);
	void LogLPenHelperEnd(int result);

	void RegisterFunction(LogBuffer& buffer);
	const UDFMap& GetUDFMap() const { return m_udfMap; }
	void** LogUdfArgument(void* key, void** lpArgument);
	void LogUdfEnd(void* key, XlFuncResult result);

protected:
	void GetXlFunctionName(int xlfn, std::wstring& str);
	void GetXlFunctionTypeStr(int xlfn, std::wstring& str);
	void GetXlResultName(XLCALL_RESULT res, std::wstring& str);
	void GetXloperTypeName(int type, std::wstring& str);
	void GetXloperErrName(XLOPER_ERRTYPE type, std::wstring& str);
	void GetPascalString(LPCSTR, std::wstring& result);
	void GetPascalString(LPCWSTR, std::wstring& result);
	BOOL LogHelper::WStrToStr(const std::wstring& wstr, std::string& str);

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
	void PrintTopBuffer(UINT deep);
	void PrintEnterRow(const std::wstring& name);
	void PrintLeaveRow(const std::wstring& name);

	HRESULT ParseArgumentType(LPCWSTR lpArgType, XllFuncInfo& info);
	XlCallArgType ParseVoidRet(WCHAR typeChar);
	void GetUDFArgTypeName(XlCallArgType type, std::wstring& name);
	DWORD GetUDFArgValue(XlCallArgType type, void** lparray, std::wstring& name);

private:
	UINT64 m_nLineCount;
	UINT m_nArrayCount;
	std::mutex	m_logFileMutex;
	std::mutex	m_arrayMutex;
	int m_bPause;
	bool m_bFirstLog;

	std::wstring m_sLogPath;
	std::wstring m_sLogFile;

	std::wofstream m_fileStream;
	UDFMap m_udfMap;

	ShellCode* m_codes;
	UINT m_nCodePos;
	std::vector<LogBuffer> m_callstack;

	static const UINT nMaxUDFuncNum = 5000;
	static LogHelper g_Instance;
};

extern DWORD PASCAL UDFHook();
#include "loghelper.inl"

#endif