#ifndef __EXCEL_PROCXY__
#define __EXCEL_PROCXY__

#include <ctype.h>
#include <windows.h>
#include "xlcall.h"
#include <framewrk.h>

#define Excel_MdCallBack "MdCallBack"
#define Excel_MdCallBack12 "MdCallBack12"
#define Excel_LPenHelper "_LPenHelper"

typedef int (PASCAL *ProcMdCallBack) (int xlfn, int coper, LPXLOPER *rgpxloper, LPXLOPER xloperRes);
typedef int (PASCAL *ProcMdCallBack12) (int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res);
typedef int (PASCAL *Proc_LPenHelper) (int wCode, void* lpv);

extern ProcMdCallBack Real_MdCallBack;
extern ProcMdCallBack12 Real_MdCallBack12;
extern Proc_LPenHelper Real__LPenHelper;

int PASCAL Mine_MdCallBack(int xlfn, int coper, LPXLOPER *rgpxloper, LPXLOPER xloperRes);
int PASCAL Mine_MdCallBack12(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res);
int PASCAL Mine__LPenHelper(int wCode, void* lpv);

#endif