#include "ExcelProcxy.h"
#include "loghelper.h"

int PASCAL Mine_MdCallBack(int xlfn, int coper, LPXLOPER *rgpxloper, LPXLOPER xloperRes)
{
	LogBuffer buffer;

	LogHelper::Instance().LogFunctionBegin(xlfn, coper, rgpxloper, buffer);
	int res = Real_MdCallBack(xlfn, coper, rgpxloper, xloperRes);
	LogHelper::Instance().LogFunctionEnd(res, xloperRes, buffer);

	return res;
}

int PASCAL Mine_MdCallBack12(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res)
{
	LogBuffer buffer;

	LogHelper::Instance().LogFunctionBegin(xlfn, coper, rgpxloper12, buffer);
	int res = Real_MdCallBack12(xlfn, coper, rgpxloper12, xloper12Res);
	LogHelper::Instance().LogFunctionEnd(res, xloper12Res, buffer);

	return res;
}

int PASCAL Mine__LPenHelper(int wCode, void* lpv)
{
	LogBuffer buffer;

	LogHelper::Instance().LogLPenHelperBegin(wCode, lpv, buffer);
	int res = Real__LPenHelper(wCode, lpv);
	LogHelper::Instance().LogLPenHelperEnd(res, buffer);

	return res;
}
