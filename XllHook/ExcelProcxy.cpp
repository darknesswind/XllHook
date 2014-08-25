#include "ExcelProcxy.h"
#include "loghelper.h"

int PASCAL Mine_MdCallBack(int xlfn, int coper, LPXLOPER *rgpxloper, LPXLOPER xloperRes)
{
	XLOPER oper = { 0, xltypeMissing };

	LogHelper::Instance().LogFunctionBegin(xlfn, coper, rgpxloper);
	if (!xloperRes)
		xloperRes = &oper;

	int res = Real_MdCallBack(xlfn, coper, rgpxloper, xloperRes);
	LogHelper::Instance().LogFunctionEnd(res, xloperRes);

	LPXLOPER lpTmp = &oper;
	Real_MdCallBack(xlcall::xlFree, 1, &lpTmp, NULL);
	return res;
}

int PASCAL Mine_MdCallBack12(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res)
{
	XLOPER12 oper = { 0, xltypeMissing };

	LogHelper::Instance().LogFunctionBegin(xlfn, coper, rgpxloper12);
	if (!xloper12Res)
		xloper12Res = &oper;

	int res = Real_MdCallBack12(xlfn, coper, rgpxloper12, xloper12Res);

	LogHelper::Instance().LogFunctionEnd(res, xloper12Res);

	LPXLOPER12 lpTmp = &oper;
	Real_MdCallBack12(xlcall::xlFree, 1, &lpTmp, NULL);
	return res;
}

int PASCAL Mine__LPenHelper(int wCode, void* lpv)
{
	LogHelper::Instance().LogLPenHelperBegin(wCode, lpv);
	int res = Real__LPenHelper(wCode, lpv);
	LogHelper::Instance().LogLPenHelperEnd(res);

	return res;
}
