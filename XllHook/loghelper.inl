#include <sstream>
#include <iomanip>

template <class LPOperType>
void LogHelper::LogFunctionBegin(int xlfn, int coper, LPOperType *rgpxloper, LogBuffer& buffer)
{
	if (m_bPause)
		return;

	GetXlFunctionTypeStr(xlfn, buffer.sFuncAttr);
	GetXlFunctionName(xlfn, buffer.sFuncName);

	int nFunc = FuncNumMask & xlfn;
	if (rgpxloper)
	{
		for (int i = 0; i < coper; ++i)
		{
			buffer.argsOperType.push_back(std::wstring());
			buffer.argsOperValue.push_back(std::wstring());
			if (!rgpxloper[i])
				continue;

			LogXloper(rgpxloper[i], buffer.argsOperType.back(), buffer.argsOperValue.back());
		}

		if (xlcall::xlCoerce == nFunc && buffer.argsOperValue.size() >= 2)
		{
			int nOperType = _wtoi(buffer.argsOperValue[1].c_str());
			GetXloperTypeName(nOperType, buffer.argsOperValue[1]);
		}
	}
}

template <class LPOperType>
void LogHelper::LogFunctionEnd(int result, LPOperType xloperRes, LogBuffer& buffer)
{
	if (m_bPause)
		return;

	if (xloperRes)
	{
		LogXloper(xloperRes, buffer.sResOperType, buffer.sResOperValue);
	}

	GetXlResultName((XLCALL_RESULT)result, buffer.sResult);
	PrintBuffer(buffer);
}

template <class LPOperType>
void LogHelper::LogXloper(LPOperType lpOper, std::wstring& sType, std::wstring& sVal)
{
	if (!lpOper)
		return;

	GetXloperTypeName(lpOper->xltype, sType);

	std::wstringstream stream;

	XLOPERTYPE type = (XLOPERTYPE)(XLOPER_TYPEMASK & lpOper->xltype);
	switch (type)
	{
	case xltypeNum:
		stream << std::setprecision(12) << lpOper->val.num;
		break;
	case xltypeStr:
		GetPascalString(lpOper->val.str, sVal);
		break;
	case xltypeBool:
		stream << (lpOper->val.xbool ? __X("True") : __X("False"));
		break;
	case xltypeRef:
		LogReferences(lpOper->val.mref.idSheet, lpOper->val.mref.lpmref, stream);
		break;
	case xltypeErr:
		GetXloperErrName((XLOPER_ERRTYPE)lpOper->val.err, sVal);
		break;
	case xltypeFlow:
		LogXloperFlow(lpOper, stream);
		break;
	case xltypeMulti:
		LogOperArray(
			lpOper->val.array.rows,
			lpOper->val.array.columns,
			lpOper->val.array.lparray,
			stream);
		break;
	case xltypeMissing:
		break;
	case xltypeNil:
		break;
	case xltypeSRef:
		LogSingleRef(lpOper->val.sref.ref, stream);
		break;
	case xltypeInt:
		stream << lpOper->val.w;
		break;
	case xltypeBigData:
		break;
	default:
		stream << __X("Unknown");
		break;
	}

	if (!stream.str().empty())
	{
		sVal = stream.str();
	}
}

template <class XLRefType>
void LogHelper::LogSingleRef(const XLRefType& sref, std::wstringstream& stream)
{
	stream << __Xc('R') << sref.rwFirst + 1
		<< __Xc('C') << sref.colFirst + 1;

	if (sref.rwFirst != sref.rwLast || sref.colFirst != sref.colLast)
	{
		stream << __X(":R")	<< sref.rwLast + 1
			<< __Xc('C') << sref.colLast + 1;
	}
}

template <class XLMRefType>
void LogHelper::LogReferences(IDSHEET iSheet, const XLMRefType* lpmref, std::wstringstream& stream)
{
	stream << __X("Sheet(") << std::hex << iSheet << __X(")!");
	if (lpmref)
	{
		for (int i = 0; i < lpmref->count; ++i)
		{
			stream << __X(" ");
			LogSingleRef(lpmref->reftbl[i], stream);
		}
	}
}

template <class LPOperType>
void LogHelper::LogXloperFlow(LPOperType lpOper, std::wstringstream& stream)
{
	if (!lpOper)
		return;

	switch (lpOper->val.flow.xlflow)
	{
	case xlflowHalt:
		stream << __X("Halt");
		break;
	case xlflowGoto:
		stream << __X("Goto Sheet(") << std::hex << lpOper->val.flow.valflow.idSheet
			<< __Xc(")!R") << lpOper->val.flow.rw + 1
			<< __Xc('C') << lpOper->val.flow.col + 1;
		break;
	case xlflowRestart:
		stream << __X("Restart ") << lpOper->val.flow.valflow.level;
		break;
	case xlflowPause:
		stream << __X("Pause ") << lpOper->val.flow.valflow.tbctrl;
		break;
	case xlflowResume:
		stream << __X("Resume");
		break;
	default:
		stream << __X("UnknownType: ") << lpOper->val.flow.xlflow;
	}
}

template <class LPOperType>
void LogHelper::LogOperArray(RW row, COL col, LPOperType lpArray, std::wstringstream& stream)
{
	stream << row << __Xc('x') << col;
	if (!lpArray || row <= 0 || col <= 0)
		return;

	UINT64 nSize = row * col;
	if (nSize > 500)
		return;

	if (nSize <= 5)
	{
		for (UINT i = 0; i < nSize; ++i)
		{
			std::wstring sType;
			std::wstring sVal;
			LogXloper(&lpArray[i], sType, sVal);
			stream << sType << __X(" {") << sVal << __Xc('}');
		}
	}
	else
	{
		m_arrayMutex.lock();
		++m_nArrayCount;
		m_arrayMutex.unlock();

		stream << __X(" in Array") << m_nArrayCount;
		LogArrayToFile(row, col, lpArray, m_nArrayCount);
	}
}

template <class LPOperType>
void LogHelper::LogArrayToFile(RW row, COL col, LPOperType lpArray, UINT id)
{
	if (!lpArray)
		return;

	std::wfstream arrayStream;
	{
		WCHAR fullPath[MAX_PATH];
		swprintf_s(fullPath, __X("%s\\array%i.csv"), m_sLogPath.c_str(), id);

		arrayStream.open(fullPath, std::wfstream::out);
	}
	if (!arrayStream.is_open())
		return;

	arrayStream.imbue(std::locale(""));

	UINT nPos = 0;
	for (RW r = 0; r < row; ++r)
	{
		for (COL c = 0; c < col; ++c)
		{
			std::wstring sType;
			std::wstring sVal;
			LogXloper(&lpArray[nPos], sType, sVal);
			if (c > 0)
				arrayStream << __Xc(',');

			arrayStream << sType << __Xc(',') << sVal;

			if (1 == row)
				arrayStream << std::endl;

			++nPos;
		}
		arrayStream << std::endl;
		arrayStream.flush();
	}
	arrayStream.close();
}
