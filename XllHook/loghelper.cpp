#include "loghelper.h"
#include <ShlObj.h>
#include <ctime>
#include <cassert>
#include "XllHook.h"

LogHelper LogHelper::g_Instance;

#define EnumNameXlCase(value)	\
	EnumNameCase2(xlcall::, value)
#define EnumNameCase(value)	\
	case value: stream << STRINGIFY(value); break
#define EnumNameCase2(prefix, value)	\
	case prefix##value: stream << STRINGIFY(value); break

LogHelper::LogHelper()
	: m_nLineCount(0)
	, m_nArrayCount(0)
	, m_bPause(0)
	, m_bFirstLog(true)
	, m_bOpened(false)
{
	m_nCodePos = 0;
	m_codes = (ShellCode*)VirtualAlloc(NULL,
		nMaxUDFuncNum * sizeof(ShellCode),
		MEM_COMMIT,
		PAGE_EXECUTE_READWRITE);
}

LogHelper::~LogHelper()
{
	VirtualFree(m_codes, 0, MEM_RELEASE);
}

void LogHelper::GetXlFunctionName(int xlfn, std::wstring& str)
{
	std::wstringstream stream;
	int nFunc = FuncNumMask & xlfn;
	switch (nFunc)
	{
		// Special Function
#if 1
		EnumNameXlCase(xlFree);
		EnumNameXlCase(xlStack);
		EnumNameXlCase(xlCoerce);
		EnumNameXlCase(xlSet);
		EnumNameXlCase(xlSheetId);
		EnumNameXlCase(xlSheetNm);
		EnumNameXlCase(xlAbort);
		EnumNameXlCase(xlGetInst);
		EnumNameXlCase(xlGetHwnd);
		EnumNameXlCase(xlGetName);
		EnumNameXlCase(xlEnableXLMsgs);
		EnumNameXlCase(xlDisableXLMsgs);
		EnumNameXlCase(xlDefineBinaryName);
		EnumNameXlCase(xlGetBinaryName);
		EnumNameXlCase(xlGetFmlaInfo);
		EnumNameXlCase(xlGetMouseInfo);
		EnumNameXlCase(xlAsyncReturn);
		EnumNameXlCase(xlEventRegister);
		EnumNameXlCase(xlRunningOnCluster);
		EnumNameXlCase(xlGetInstPtr);
#endif

		// Normal Function
#if 1
		EnumNameXlCase(xlfCount);
		EnumNameXlCase(xlfIsna);
		EnumNameXlCase(xlfIserror);
		EnumNameXlCase(xlfSum);
		EnumNameXlCase(xlfAverage);
		EnumNameXlCase(xlfMin);
		EnumNameXlCase(xlfMax);
		EnumNameXlCase(xlfRow);
		EnumNameXlCase(xlfColumn);
		EnumNameXlCase(xlfNa);
		EnumNameXlCase(xlfNpv);
		EnumNameXlCase(xlfStdev);
		EnumNameXlCase(xlfDollar);
		EnumNameXlCase(xlfFixed);
		EnumNameXlCase(xlfSin);
		EnumNameXlCase(xlfCos);
		EnumNameXlCase(xlfTan);
		EnumNameXlCase(xlfAtan);
		EnumNameXlCase(xlfPi);
		EnumNameXlCase(xlfSqrt);
		EnumNameXlCase(xlfExp);
		EnumNameXlCase(xlfLn);
		EnumNameXlCase(xlfLog10);
		EnumNameXlCase(xlfAbs);
		EnumNameXlCase(xlfInt);
		EnumNameXlCase(xlfSign);
		EnumNameXlCase(xlfRound);
		EnumNameXlCase(xlfLookup);
		EnumNameXlCase(xlfIndex);
		EnumNameXlCase(xlfRept);
		EnumNameXlCase(xlfMid);
		EnumNameXlCase(xlfLen);
		EnumNameXlCase(xlfValue);
		EnumNameXlCase(xlfTrue);
		EnumNameXlCase(xlfFalse);
		EnumNameXlCase(xlfAnd);
		EnumNameXlCase(xlfOr);
		EnumNameXlCase(xlfNot);
		EnumNameXlCase(xlfMod);
		EnumNameXlCase(xlfDcount);
		EnumNameXlCase(xlfDsum);
		EnumNameXlCase(xlfDaverage);
		EnumNameXlCase(xlfDmin);
		EnumNameXlCase(xlfDmax);
		EnumNameXlCase(xlfDstdev);
		EnumNameXlCase(xlfVar);
		EnumNameXlCase(xlfDvar);
		EnumNameXlCase(xlfText);
		EnumNameXlCase(xlfLinest);
		EnumNameXlCase(xlfTrend);
		EnumNameXlCase(xlfLogest);
		EnumNameXlCase(xlfGrowth);
		EnumNameXlCase(xlfGoto);
		EnumNameXlCase(xlfHalt);
		EnumNameXlCase(xlfPv);
		EnumNameXlCase(xlfFv);
		EnumNameXlCase(xlfNper);
		EnumNameXlCase(xlfPmt);
		EnumNameXlCase(xlfRate);
		EnumNameXlCase(xlfMirr);
		EnumNameXlCase(xlfIrr);
		EnumNameXlCase(xlfRand);
		EnumNameXlCase(xlfMatch);
		EnumNameXlCase(xlfDate);
		EnumNameXlCase(xlfTime);
		EnumNameXlCase(xlfDay);
		EnumNameXlCase(xlfMonth);
		EnumNameXlCase(xlfYear);
		EnumNameXlCase(xlfWeekday);
		EnumNameXlCase(xlfHour);
		EnumNameXlCase(xlfMinute);
		EnumNameXlCase(xlfSecond);
		EnumNameXlCase(xlfNow);
		EnumNameXlCase(xlfAreas);
		EnumNameXlCase(xlfRows);
		EnumNameXlCase(xlfColumns);
		EnumNameXlCase(xlfOffset);
		EnumNameXlCase(xlfAbsref);
		EnumNameXlCase(xlfRelref);
		EnumNameXlCase(xlfArgument);
		EnumNameXlCase(xlfSearch);
		EnumNameXlCase(xlfTranspose);
		EnumNameXlCase(xlfError);
		EnumNameXlCase(xlfStep);
		EnumNameXlCase(xlfType);
		EnumNameXlCase(xlfEcho);
		EnumNameXlCase(xlfSetName);
		EnumNameXlCase(xlfCaller);
		EnumNameXlCase(xlfDeref);
		EnumNameXlCase(xlfWindows);
		EnumNameXlCase(xlfSeries);
		EnumNameXlCase(xlfDocuments);
		EnumNameXlCase(xlfActiveCell);
		EnumNameXlCase(xlfSelection);
		EnumNameXlCase(xlfResult);
		EnumNameXlCase(xlfAtan2);
		EnumNameXlCase(xlfAsin);
		EnumNameXlCase(xlfAcos);
		EnumNameXlCase(xlfChoose);
		EnumNameXlCase(xlfHlookup);
		EnumNameXlCase(xlfVlookup);
		EnumNameXlCase(xlfLinks);
		EnumNameXlCase(xlfInput);
		EnumNameXlCase(xlfIsref);
		EnumNameXlCase(xlfGetFormula);
		EnumNameXlCase(xlfGetName);
		EnumNameXlCase(xlfSetValue);
		EnumNameXlCase(xlfLog);
		EnumNameXlCase(xlfExec);
		EnumNameXlCase(xlfChar);
		EnumNameXlCase(xlfLower);
		EnumNameXlCase(xlfUpper);
		EnumNameXlCase(xlfProper);
		EnumNameXlCase(xlfLeft);
		EnumNameXlCase(xlfRight);
		EnumNameXlCase(xlfExact);
		EnumNameXlCase(xlfTrim);
		EnumNameXlCase(xlfReplace);
		EnumNameXlCase(xlfSubstitute);
		EnumNameXlCase(xlfCode);
		EnumNameXlCase(xlfNames);
		EnumNameXlCase(xlfDirectory);
		EnumNameXlCase(xlfFind);
		EnumNameXlCase(xlfCell);
		EnumNameXlCase(xlfIserr);
		EnumNameXlCase(xlfIstext);
		EnumNameXlCase(xlfIsnumber);
		EnumNameXlCase(xlfIsblank);
		EnumNameXlCase(xlfT);
		EnumNameXlCase(xlfN);
		EnumNameXlCase(xlfFopen);
		EnumNameXlCase(xlfFclose);
		EnumNameXlCase(xlfFsize);
		EnumNameXlCase(xlfFreadln);
		EnumNameXlCase(xlfFread);
		EnumNameXlCase(xlfFwriteln);
		EnumNameXlCase(xlfFwrite);
		EnumNameXlCase(xlfFpos);
		EnumNameXlCase(xlfDatevalue);
		EnumNameXlCase(xlfTimevalue);
		EnumNameXlCase(xlfSln);
		EnumNameXlCase(xlfSyd);
		EnumNameXlCase(xlfDdb);
		EnumNameXlCase(xlfGetDef);
		EnumNameXlCase(xlfReftext);
		EnumNameXlCase(xlfTextref);
		EnumNameXlCase(xlfIndirect);
		EnumNameXlCase(xlfRegister);
		EnumNameXlCase(xlfCall);
		EnumNameXlCase(xlfAddBar);
		EnumNameXlCase(xlfAddMenu);
		EnumNameXlCase(xlfAddCommand);
		EnumNameXlCase(xlfEnableCommand);
		EnumNameXlCase(xlfCheckCommand);
		EnumNameXlCase(xlfRenameCommand);
		EnumNameXlCase(xlfShowBar);
		EnumNameXlCase(xlfDeleteMenu);
		EnumNameXlCase(xlfDeleteCommand);
		EnumNameXlCase(xlfGetChartItem);
		EnumNameXlCase(xlfDialogBox);
		EnumNameXlCase(xlfClean);
		EnumNameXlCase(xlfMdeterm);
		EnumNameXlCase(xlfMinverse);
		EnumNameXlCase(xlfMmult);
		EnumNameXlCase(xlfFiles);
		EnumNameXlCase(xlfIpmt);
		EnumNameXlCase(xlfPpmt);
		EnumNameXlCase(xlfCounta);
		EnumNameXlCase(xlfCancelKey);
		EnumNameXlCase(xlfInitiate);
		EnumNameXlCase(xlfRequest);
		EnumNameXlCase(xlfPoke);
		EnumNameXlCase(xlfExecute);
		EnumNameXlCase(xlfTerminate);
		EnumNameXlCase(xlfRestart);
		EnumNameXlCase(xlfHelp);
		EnumNameXlCase(xlfGetBar);
		EnumNameXlCase(xlfProduct);
		EnumNameXlCase(xlfFact);
		EnumNameXlCase(xlfGetCell);
		EnumNameXlCase(xlfGetWorkspace);
		EnumNameXlCase(xlfGetWindow);
		EnumNameXlCase(xlfGetDocument);
		EnumNameXlCase(xlfDproduct);
		EnumNameXlCase(xlfIsnontext);
		EnumNameXlCase(xlfGetNote);
		EnumNameXlCase(xlfNote);
		EnumNameXlCase(xlfStdevp);
		EnumNameXlCase(xlfVarp);
		EnumNameXlCase(xlfDstdevp);
		EnumNameXlCase(xlfDvarp);
		EnumNameXlCase(xlfTrunc);
		EnumNameXlCase(xlfIslogical);
		EnumNameXlCase(xlfDcounta);
		EnumNameXlCase(xlfDeleteBar);
		EnumNameXlCase(xlfUnregister);
		EnumNameXlCase(xlfUsdollar);
		EnumNameXlCase(xlfFindb);
		EnumNameXlCase(xlfSearchb);
		EnumNameXlCase(xlfReplaceb);
		EnumNameXlCase(xlfLeftb);
		EnumNameXlCase(xlfRightb);
		EnumNameXlCase(xlfMidb);
		EnumNameXlCase(xlfLenb);
		EnumNameXlCase(xlfRoundup);
		EnumNameXlCase(xlfRounddown);
		EnumNameXlCase(xlfAsc);
		EnumNameXlCase(xlfDbcs);
		EnumNameXlCase(xlfRank);
		EnumNameXlCase(xlfAddress);
		EnumNameXlCase(xlfDays360);
		EnumNameXlCase(xlfToday);
		EnumNameXlCase(xlfVdb);
		EnumNameXlCase(xlfMedian);
		EnumNameXlCase(xlfSumproduct);
		EnumNameXlCase(xlfSinh);
		EnumNameXlCase(xlfCosh);
		EnumNameXlCase(xlfTanh);
		EnumNameXlCase(xlfAsinh);
		EnumNameXlCase(xlfAcosh);
		EnumNameXlCase(xlfAtanh);
		EnumNameXlCase(xlfDget);
		EnumNameXlCase(xlfCreateObject);
		EnumNameXlCase(xlfVolatile);
		EnumNameXlCase(xlfLastError);
		EnumNameXlCase(xlfCustomUndo);
		EnumNameXlCase(xlfCustomRepeat);
		EnumNameXlCase(xlfFormulaConvert);
		EnumNameXlCase(xlfGetLinkInfo);
		EnumNameXlCase(xlfTextBox);
		EnumNameXlCase(xlfInfo);
		EnumNameXlCase(xlfGroup);
		EnumNameXlCase(xlfGetObject);
		EnumNameXlCase(xlfDb);
		EnumNameXlCase(xlfPause);
		EnumNameXlCase(xlfResume);
		EnumNameXlCase(xlfFrequency);
		EnumNameXlCase(xlfAddToolbar);
		EnumNameXlCase(xlfDeleteToolbar);
		EnumNameXlCase(xlUDF);
		EnumNameXlCase(xlfResetToolbar);
		EnumNameXlCase(xlfEvaluate);
		EnumNameXlCase(xlfGetToolbar);
		EnumNameXlCase(xlfGetTool);
		EnumNameXlCase(xlfSpellingCheck);
		EnumNameXlCase(xlfErrorType);
		EnumNameXlCase(xlfAppTitle);
		EnumNameXlCase(xlfWindowTitle);
		EnumNameXlCase(xlfSaveToolbar);
		EnumNameXlCase(xlfEnableTool);
		EnumNameXlCase(xlfPressTool);
		EnumNameXlCase(xlfRegisterId);
		EnumNameXlCase(xlfGetWorkbook);
		EnumNameXlCase(xlfAvedev);
		EnumNameXlCase(xlfBetadist);
		EnumNameXlCase(xlfGammaln);
		EnumNameXlCase(xlfBetainv);
		EnumNameXlCase(xlfBinomdist);
		EnumNameXlCase(xlfChidist);
		EnumNameXlCase(xlfChiinv);
		EnumNameXlCase(xlfCombin);
		EnumNameXlCase(xlfConfidence);
		EnumNameXlCase(xlfCritbinom);
		EnumNameXlCase(xlfEven);
		EnumNameXlCase(xlfExpondist);
		EnumNameXlCase(xlfFdist);
		EnumNameXlCase(xlfFinv);
		EnumNameXlCase(xlfFisher);
		EnumNameXlCase(xlfFisherinv);
		EnumNameXlCase(xlfFloor);
		EnumNameXlCase(xlfGammadist);
		EnumNameXlCase(xlfGammainv);
		EnumNameXlCase(xlfCeiling);
		EnumNameXlCase(xlfHypgeomdist);
		EnumNameXlCase(xlfLognormdist);
		EnumNameXlCase(xlfLoginv);
		EnumNameXlCase(xlfNegbinomdist);
		EnumNameXlCase(xlfNormdist);
		EnumNameXlCase(xlfNormsdist);
		EnumNameXlCase(xlfNorminv);
		EnumNameXlCase(xlfNormsinv);
		EnumNameXlCase(xlfStandardize);
		EnumNameXlCase(xlfOdd);
		EnumNameXlCase(xlfPermut);
		EnumNameXlCase(xlfPoisson);
		EnumNameXlCase(xlfTdist);
		EnumNameXlCase(xlfWeibull);
		EnumNameXlCase(xlfSumxmy2);
		EnumNameXlCase(xlfSumx2my2);
		EnumNameXlCase(xlfSumx2py2);
		EnumNameXlCase(xlfChitest);
		EnumNameXlCase(xlfCorrel);
		EnumNameXlCase(xlfCovar);
		EnumNameXlCase(xlfForecast);
		EnumNameXlCase(xlfFtest);
		EnumNameXlCase(xlfIntercept);
		EnumNameXlCase(xlfPearson);
		EnumNameXlCase(xlfRsq);
		EnumNameXlCase(xlfSteyx);
		EnumNameXlCase(xlfSlope);
		EnumNameXlCase(xlfTtest);
		EnumNameXlCase(xlfProb);
		EnumNameXlCase(xlfDevsq);
		EnumNameXlCase(xlfGeomean);
		EnumNameXlCase(xlfHarmean);
		EnumNameXlCase(xlfSumsq);
		EnumNameXlCase(xlfKurt);
		EnumNameXlCase(xlfSkew);
		EnumNameXlCase(xlfZtest);
		EnumNameXlCase(xlfLarge);
		EnumNameXlCase(xlfSmall);
		EnumNameXlCase(xlfQuartile);
		EnumNameXlCase(xlfPercentile);
		EnumNameXlCase(xlfPercentrank);
		EnumNameXlCase(xlfMode);
		EnumNameXlCase(xlfTrimmean);
		EnumNameXlCase(xlfTinv);
		EnumNameXlCase(xlfMovieCommand);
		EnumNameXlCase(xlfGetMovie);
		EnumNameXlCase(xlfConcatenate);
		EnumNameXlCase(xlfPower);
		EnumNameXlCase(xlfPivotAddData);
		EnumNameXlCase(xlfGetPivotTable);
		EnumNameXlCase(xlfGetPivotField);
		EnumNameXlCase(xlfGetPivotItem);
		EnumNameXlCase(xlfRadians);
		EnumNameXlCase(xlfDegrees);
		EnumNameXlCase(xlfSubtotal);
		EnumNameXlCase(xlfSumif);
		EnumNameXlCase(xlfCountif);
		EnumNameXlCase(xlfCountblank);
		EnumNameXlCase(xlfScenarioGet);
		EnumNameXlCase(xlfOptionsListsGet);
		EnumNameXlCase(xlfIspmt);
		EnumNameXlCase(xlfDatedif);
		EnumNameXlCase(xlfDatestring);
		EnumNameXlCase(xlfNumberstring);
		EnumNameXlCase(xlfRoman);
		EnumNameXlCase(xlfOpenDialog);
		EnumNameXlCase(xlfSaveDialog);
		EnumNameXlCase(xlfViewGet);
		EnumNameXlCase(xlfGetpivotdata);
		EnumNameXlCase(xlfHyperlink);
		EnumNameXlCase(xlfPhonetic);
		EnumNameXlCase(xlfAveragea);
		EnumNameXlCase(xlfMaxa);
		EnumNameXlCase(xlfMina);
		EnumNameXlCase(xlfStdevpa);
		EnumNameXlCase(xlfVarpa);
		EnumNameXlCase(xlfStdeva);
		EnumNameXlCase(xlfVara);
		EnumNameXlCase(xlfBahttext);
		EnumNameXlCase(xlfThaidayofweek);
		EnumNameXlCase(xlfThaidigit);
		EnumNameXlCase(xlfThaimonthofyear);
		EnumNameXlCase(xlfThainumsound);
		EnumNameXlCase(xlfThainumstring);
		EnumNameXlCase(xlfThaistringlength);
		EnumNameXlCase(xlfIsthaidigit);
		EnumNameXlCase(xlfRoundbahtdown);
		EnumNameXlCase(xlfRoundbahtup);
		EnumNameXlCase(xlfThaiyear);
		EnumNameXlCase(xlfRtd);
		EnumNameXlCase(xlfCubevalue);
		EnumNameXlCase(xlfCubemember);
		EnumNameXlCase(xlfCubememberproperty);
		EnumNameXlCase(xlfCuberankedmember);
		EnumNameXlCase(xlfHex2bin);
		EnumNameXlCase(xlfHex2dec);
		EnumNameXlCase(xlfHex2oct);
		EnumNameXlCase(xlfDec2bin);
		EnumNameXlCase(xlfDec2hex);
		EnumNameXlCase(xlfDec2oct);
		EnumNameXlCase(xlfOct2bin);
		EnumNameXlCase(xlfOct2hex);
		EnumNameXlCase(xlfOct2dec);
		EnumNameXlCase(xlfBin2dec);
		EnumNameXlCase(xlfBin2oct);
		EnumNameXlCase(xlfBin2hex);
		EnumNameXlCase(xlfImsub);
		EnumNameXlCase(xlfImdiv);
		EnumNameXlCase(xlfImpower);
		EnumNameXlCase(xlfImabs);
		EnumNameXlCase(xlfImsqrt);
		EnumNameXlCase(xlfImln);
		EnumNameXlCase(xlfImlog2);
		EnumNameXlCase(xlfImlog10);
		EnumNameXlCase(xlfImsin);
		EnumNameXlCase(xlfImcos);
		EnumNameXlCase(xlfImexp);
		EnumNameXlCase(xlfImargument);
		EnumNameXlCase(xlfImconjugate);
		EnumNameXlCase(xlfImaginary);
		EnumNameXlCase(xlfImreal);
		EnumNameXlCase(xlfComplex);
		EnumNameXlCase(xlfImsum);
		EnumNameXlCase(xlfImproduct);
		EnumNameXlCase(xlfSeriessum);
		EnumNameXlCase(xlfFactdouble);
		EnumNameXlCase(xlfSqrtpi);
		EnumNameXlCase(xlfQuotient);
		EnumNameXlCase(xlfDelta);
		EnumNameXlCase(xlfGestep);
		EnumNameXlCase(xlfIseven);
		EnumNameXlCase(xlfIsodd);
		EnumNameXlCase(xlfMround);
		EnumNameXlCase(xlfErf);
		EnumNameXlCase(xlfErfc);
		EnumNameXlCase(xlfBesselj);
		EnumNameXlCase(xlfBesselk);
		EnumNameXlCase(xlfBessely);
		EnumNameXlCase(xlfBesseli);
		EnumNameXlCase(xlfXirr);
		EnumNameXlCase(xlfXnpv);
		EnumNameXlCase(xlfPricemat);
		EnumNameXlCase(xlfYieldmat);
		EnumNameXlCase(xlfIntrate);
		EnumNameXlCase(xlfReceived);
		EnumNameXlCase(xlfDisc);
		EnumNameXlCase(xlfPricedisc);
		EnumNameXlCase(xlfYielddisc);
		EnumNameXlCase(xlfTbilleq);
		EnumNameXlCase(xlfTbillprice);
		EnumNameXlCase(xlfTbillyield);
		EnumNameXlCase(xlfPrice);
		EnumNameXlCase(xlfYield);
		EnumNameXlCase(xlfDollarde);
		EnumNameXlCase(xlfDollarfr);
		EnumNameXlCase(xlfNominal);
		EnumNameXlCase(xlfEffect);
		EnumNameXlCase(xlfCumprinc);
		EnumNameXlCase(xlfCumipmt);
		EnumNameXlCase(xlfEdate);
		EnumNameXlCase(xlfEomonth);
		EnumNameXlCase(xlfYearfrac);
		EnumNameXlCase(xlfCoupdaybs);
		EnumNameXlCase(xlfCoupdays);
		EnumNameXlCase(xlfCoupdaysnc);
		EnumNameXlCase(xlfCoupncd);
		EnumNameXlCase(xlfCoupnum);
		EnumNameXlCase(xlfCouppcd);
		EnumNameXlCase(xlfDuration);
		EnumNameXlCase(xlfMduration);
		EnumNameXlCase(xlfOddlprice);
		EnumNameXlCase(xlfOddlyield);
		EnumNameXlCase(xlfOddfprice);
		EnumNameXlCase(xlfOddfyield);
		EnumNameXlCase(xlfRandbetween);
		EnumNameXlCase(xlfWeeknum);
		EnumNameXlCase(xlfAmordegrc);
		EnumNameXlCase(xlfAmorlinc);
		EnumNameXlCase(xlfConvert);
		EnumNameXlCase(xlfAccrint);
		EnumNameXlCase(xlfAccrintm);
		EnumNameXlCase(xlfWorkday);
		EnumNameXlCase(xlfNetworkdays);
		EnumNameXlCase(xlfGcd);
		EnumNameXlCase(xlfMultinomial);
		EnumNameXlCase(xlfLcm);
		EnumNameXlCase(xlfFvschedule);
		EnumNameXlCase(xlfCubekpimember);
		EnumNameXlCase(xlfCubeset);
		EnumNameXlCase(xlfCubesetcount);
		EnumNameXlCase(xlfIferror);
		EnumNameXlCase(xlfCountifs);
		EnumNameXlCase(xlfSumifs);
		EnumNameXlCase(xlfAverageif);
		EnumNameXlCase(xlfAverageifs);
		EnumNameXlCase(xlfAggregate);
		EnumNameXlCase(xlfBinom_dist);
		EnumNameXlCase(xlfBinom_inv);
		EnumNameXlCase(xlfConfidence_norm);
		EnumNameXlCase(xlfConfidence_t);
		EnumNameXlCase(xlfChisq_test);
		EnumNameXlCase(xlfF_test);
		EnumNameXlCase(xlfCovariance_p);
		EnumNameXlCase(xlfCovariance_s);
		EnumNameXlCase(xlfExpon_dist);
		EnumNameXlCase(xlfGamma_dist);
		EnumNameXlCase(xlfGamma_inv);
		EnumNameXlCase(xlfMode_mult);
		EnumNameXlCase(xlfMode_sngl);
		EnumNameXlCase(xlfNorm_dist);
		EnumNameXlCase(xlfNorm_inv);
		EnumNameXlCase(xlfPercentile_exc);
		EnumNameXlCase(xlfPercentile_inc);
		EnumNameXlCase(xlfPercentrank_exc);
		EnumNameXlCase(xlfPercentrank_inc);
		EnumNameXlCase(xlfPoisson_dist);
		EnumNameXlCase(xlfQuartile_exc);
		EnumNameXlCase(xlfQuartile_inc);
		EnumNameXlCase(xlfRank_avg);
		EnumNameXlCase(xlfRank_eq);
		EnumNameXlCase(xlfStdev_s);
		EnumNameXlCase(xlfStdev_p);
		EnumNameXlCase(xlfT_dist);
		EnumNameXlCase(xlfT_dist_2t);
		EnumNameXlCase(xlfT_dist_rt);
		EnumNameXlCase(xlfT_inv);
		EnumNameXlCase(xlfT_inv_2t);
		EnumNameXlCase(xlfVar_s);
		EnumNameXlCase(xlfVar_p);
		EnumNameXlCase(xlfWeibull_dist);
		EnumNameXlCase(xlfNetworkdays_intl);
		EnumNameXlCase(xlfWorkday_intl);
		EnumNameXlCase(xlfEcma_ceiling);
		EnumNameXlCase(xlfIso_ceiling);
		EnumNameXlCase(xlfBeta_dist);
		EnumNameXlCase(xlfBeta_inv);
		EnumNameXlCase(xlfChisq_dist);
		EnumNameXlCase(xlfChisq_dist_rt);
		EnumNameXlCase(xlfChisq_inv);
		EnumNameXlCase(xlfChisq_inv_rt);
		EnumNameXlCase(xlfF_dist);
		EnumNameXlCase(xlfF_dist_rt);
		EnumNameXlCase(xlfF_inv);
		EnumNameXlCase(xlfF_inv_rt);
		EnumNameXlCase(xlfHypgeom_dist);
		EnumNameXlCase(xlfLognorm_dist);
		EnumNameXlCase(xlfLognorm_inv);
		EnumNameXlCase(xlfNegbinom_dist);
		EnumNameXlCase(xlfNorm_s_dist);
		EnumNameXlCase(xlfNorm_s_inv);
		EnumNameXlCase(xlfT_test);
		EnumNameXlCase(xlfZ_test);
		EnumNameXlCase(xlfErf_precise);
		EnumNameXlCase(xlfErfc_precise);
		EnumNameXlCase(xlfGammaln_precise);
		EnumNameXlCase(xlfCeiling_precise);
		EnumNameXlCase(xlfFloor_precise);
		EnumNameXlCase(xlfAcot);
		EnumNameXlCase(xlfAcoth);
		EnumNameXlCase(xlfCot);
		EnumNameXlCase(xlfCoth);
		EnumNameXlCase(xlfCsc);
		EnumNameXlCase(xlfCsch);
		EnumNameXlCase(xlfSec);
		EnumNameXlCase(xlfSech);
		EnumNameXlCase(xlfImtan);
		EnumNameXlCase(xlfImcot);
		EnumNameXlCase(xlfImcsc);
		EnumNameXlCase(xlfImcsch);
		EnumNameXlCase(xlfImsec);
		EnumNameXlCase(xlfImsech);
		EnumNameXlCase(xlfBitand);
		EnumNameXlCase(xlfBitor);
		EnumNameXlCase(xlfBitxor);
		EnumNameXlCase(xlfBitlshift);
		EnumNameXlCase(xlfBitrshift);
		EnumNameXlCase(xlfPermutationa);
		EnumNameXlCase(xlfCombina);
		EnumNameXlCase(xlfXor);
		EnumNameXlCase(xlfPduration);
		EnumNameXlCase(xlfBase);
		EnumNameXlCase(xlfDecimal);
		EnumNameXlCase(xlfDays);
		EnumNameXlCase(xlfBinom_dist_range);
		EnumNameXlCase(xlfGamma);
		EnumNameXlCase(xlfSkew_p);
		EnumNameXlCase(xlfGauss);
		EnumNameXlCase(xlfPhi);
		EnumNameXlCase(xlfRri);
		EnumNameXlCase(xlfUnichar);
		EnumNameXlCase(xlfUnicode);
		EnumNameXlCase(xlfMunit);
		EnumNameXlCase(xlfArabic);
		EnumNameXlCase(xlfIsoweeknum);
		EnumNameXlCase(xlfNumbervalue);
		EnumNameXlCase(xlfSheet);
		EnumNameXlCase(xlfSheets);
		EnumNameXlCase(xlfFormulatext);
		EnumNameXlCase(xlfIsformula);
		EnumNameXlCase(xlfIfna);
		EnumNameXlCase(xlfCeiling_math);
		EnumNameXlCase(xlfFloor_math);
		EnumNameXlCase(xlfImsinh);
		EnumNameXlCase(xlfImcosh);
		EnumNameXlCase(xlfFilterxml);
		EnumNameXlCase(xlfWebservice);
		EnumNameXlCase(xlfEncodeurl);
#endif

		// Command
#if 1
		EnumNameXlCase(xlcBeep);
		EnumNameXlCase(xlcOpen);
		EnumNameXlCase(xlcOpenLinks);
		EnumNameXlCase(xlcCloseAll);
		EnumNameXlCase(xlcSave);
		EnumNameXlCase(xlcSaveAs);
		EnumNameXlCase(xlcFileDelete);
		EnumNameXlCase(xlcPageSetup);
		EnumNameXlCase(xlcPrint);
		EnumNameXlCase(xlcPrinterSetup);
		EnumNameXlCase(xlcQuit);
		EnumNameXlCase(xlcNewWindow);
		EnumNameXlCase(xlcArrangeAll);
		EnumNameXlCase(xlcWindowSize);
		EnumNameXlCase(xlcWindowMove);
		EnumNameXlCase(xlcFull);
		EnumNameXlCase(xlcClose);
		EnumNameXlCase(xlcRun);
		EnumNameXlCase(xlcSetPrintArea);
		EnumNameXlCase(xlcSetPrintTitles);
		EnumNameXlCase(xlcSetPageBreak);
		EnumNameXlCase(xlcRemovePageBreak);
		EnumNameXlCase(xlcFont);
		EnumNameXlCase(xlcDisplay);
		EnumNameXlCase(xlcProtectDocument);
		EnumNameXlCase(xlcPrecision);
		EnumNameXlCase(xlcA1R1c1);
		EnumNameXlCase(xlcCalculateNow);
		EnumNameXlCase(xlcCalculation);
		EnumNameXlCase(xlcDataFind);
		EnumNameXlCase(xlcExtract);
		EnumNameXlCase(xlcDataDelete);
		EnumNameXlCase(xlcSetDatabase);
		EnumNameXlCase(xlcSetCriteria);
		EnumNameXlCase(xlcSort);
		EnumNameXlCase(xlcDataSeries);
		EnumNameXlCase(xlcTable);
		EnumNameXlCase(xlcFormatNumber);
		EnumNameXlCase(xlcAlignment);
		EnumNameXlCase(xlcStyle);
		EnumNameXlCase(xlcBorder);
		EnumNameXlCase(xlcCellProtection);
		EnumNameXlCase(xlcColumnWidth);
		EnumNameXlCase(xlcUndo);
		EnumNameXlCase(xlcCut);
		EnumNameXlCase(xlcCopy);
		EnumNameXlCase(xlcPaste);
		EnumNameXlCase(xlcClear);
		EnumNameXlCase(xlcPasteSpecial);
		EnumNameXlCase(xlcEditDelete);
		EnumNameXlCase(xlcInsert);
		EnumNameXlCase(xlcFillRight);
		EnumNameXlCase(xlcFillDown);
		EnumNameXlCase(xlcDefineName);
		EnumNameXlCase(xlcCreateNames);
		EnumNameXlCase(xlcFormulaGoto);
		EnumNameXlCase(xlcFormulaFind);
		EnumNameXlCase(xlcSelectLastCell);
		EnumNameXlCase(xlcShowActiveCell);
		EnumNameXlCase(xlcGalleryArea);
		EnumNameXlCase(xlcGalleryBar);
		EnumNameXlCase(xlcGalleryColumn);
		EnumNameXlCase(xlcGalleryLine);
		EnumNameXlCase(xlcGalleryPie);
		EnumNameXlCase(xlcGalleryScatter);
		EnumNameXlCase(xlcCombination);
		EnumNameXlCase(xlcPreferred);
		EnumNameXlCase(xlcAddOverlay);
		EnumNameXlCase(xlcGridlines);
		EnumNameXlCase(xlcSetPreferred);
		EnumNameXlCase(xlcAxes);
		EnumNameXlCase(xlcLegend);
		EnumNameXlCase(xlcAttachText);
		EnumNameXlCase(xlcAddArrow);
		EnumNameXlCase(xlcSelectChart);
		EnumNameXlCase(xlcSelectPlotArea);
		EnumNameXlCase(xlcPatterns);
		EnumNameXlCase(xlcMainChart);
		EnumNameXlCase(xlcOverlay);
		EnumNameXlCase(xlcScale);
		EnumNameXlCase(xlcFormatLegend);
		EnumNameXlCase(xlcFormatText);
		EnumNameXlCase(xlcEditRepeat);
		EnumNameXlCase(xlcParse);
		EnumNameXlCase(xlcJustify);
		EnumNameXlCase(xlcHide);
		EnumNameXlCase(xlcUnhide);
		EnumNameXlCase(xlcWorkspace);
		EnumNameXlCase(xlcFormula);
		EnumNameXlCase(xlcFormulaFill);
		EnumNameXlCase(xlcFormulaArray);
		EnumNameXlCase(xlcDataFindNext);
		EnumNameXlCase(xlcDataFindPrev);
		EnumNameXlCase(xlcFormulaFindNext);
		EnumNameXlCase(xlcFormulaFindPrev);
		EnumNameXlCase(xlcActivate);
		EnumNameXlCase(xlcActivateNext);
		EnumNameXlCase(xlcActivatePrev);
		EnumNameXlCase(xlcUnlockedNext);
		EnumNameXlCase(xlcUnlockedPrev);
		EnumNameXlCase(xlcCopyPicture);
		EnumNameXlCase(xlcSelect);
		EnumNameXlCase(xlcDeleteName);
		EnumNameXlCase(xlcDeleteFormat);
		EnumNameXlCase(xlcVline);
		EnumNameXlCase(xlcHline);
		EnumNameXlCase(xlcVpage);
		EnumNameXlCase(xlcHpage);
		EnumNameXlCase(xlcVscroll);
		EnumNameXlCase(xlcHscroll);
		EnumNameXlCase(xlcAlert);
		EnumNameXlCase(xlcNew);
		EnumNameXlCase(xlcCancelCopy);
		EnumNameXlCase(xlcShowClipboard);
		EnumNameXlCase(xlcMessage);
		EnumNameXlCase(xlcPasteLink);
		EnumNameXlCase(xlcAppActivate);
		EnumNameXlCase(xlcDeleteArrow);
		EnumNameXlCase(xlcRowHeight);
		EnumNameXlCase(xlcFormatMove);
		EnumNameXlCase(xlcFormatSize);
		EnumNameXlCase(xlcFormulaReplace);
		EnumNameXlCase(xlcSendKeys);
		EnumNameXlCase(xlcSelectSpecial);
		EnumNameXlCase(xlcApplyNames);
		EnumNameXlCase(xlcReplaceFont);
		EnumNameXlCase(xlcFreezePanes);
		EnumNameXlCase(xlcShowInfo);
		EnumNameXlCase(xlcSplit);
		EnumNameXlCase(xlcOnWindow);
		EnumNameXlCase(xlcOnData);
		EnumNameXlCase(xlcDisableInput);
		EnumNameXlCase(xlcEcho);
		EnumNameXlCase(xlcOutline);
		EnumNameXlCase(xlcListNames);
		EnumNameXlCase(xlcFileClose);
		EnumNameXlCase(xlcSaveWorkbook);
		EnumNameXlCase(xlcDataForm);
		EnumNameXlCase(xlcCopyChart);
		EnumNameXlCase(xlcOnTime);
		EnumNameXlCase(xlcWait);
		EnumNameXlCase(xlcFormatFont);
		EnumNameXlCase(xlcFillUp);
		EnumNameXlCase(xlcFillLeft);
		EnumNameXlCase(xlcDeleteOverlay);
		EnumNameXlCase(xlcNote);
		EnumNameXlCase(xlcShortMenus);
		EnumNameXlCase(xlcSetUpdateStatus);
		EnumNameXlCase(xlcColorPalette);
		EnumNameXlCase(xlcDeleteStyle);
		EnumNameXlCase(xlcWindowRestore);
		EnumNameXlCase(xlcWindowMaximize);
		EnumNameXlCase(xlcError);
		EnumNameXlCase(xlcChangeLink);
		EnumNameXlCase(xlcCalculateDocument);
		EnumNameXlCase(xlcOnKey);
		EnumNameXlCase(xlcAppRestore);
		EnumNameXlCase(xlcAppMove);
		EnumNameXlCase(xlcAppSize);
		EnumNameXlCase(xlcAppMinimize);
		EnumNameXlCase(xlcAppMaximize);
		EnumNameXlCase(xlcBringToFront);
		EnumNameXlCase(xlcSendToBack);
		EnumNameXlCase(xlcMainChartType);
		EnumNameXlCase(xlcOverlayChartType);
		EnumNameXlCase(xlcSelectEnd);
		EnumNameXlCase(xlcOpenMail);
		EnumNameXlCase(xlcSendMail);
		EnumNameXlCase(xlcStandardFont);
		EnumNameXlCase(xlcConsolidate);
		EnumNameXlCase(xlcSortSpecial);
		EnumNameXlCase(xlcGallery3dArea);
		EnumNameXlCase(xlcGallery3dColumn);
		EnumNameXlCase(xlcGallery3dLine);
		EnumNameXlCase(xlcGallery3dPie);
		EnumNameXlCase(xlcView3d);
		EnumNameXlCase(xlcGoalSeek);
		EnumNameXlCase(xlcWorkgroup);
		EnumNameXlCase(xlcFillGroup);
		EnumNameXlCase(xlcUpdateLink);
		EnumNameXlCase(xlcPromote);
		EnumNameXlCase(xlcDemote);
		EnumNameXlCase(xlcShowDetail);
		EnumNameXlCase(xlcUngroup);
		EnumNameXlCase(xlcObjectProperties);
		EnumNameXlCase(xlcSaveNewObject);
		EnumNameXlCase(xlcShare);
		EnumNameXlCase(xlcShareName);
		EnumNameXlCase(xlcDuplicate);
		EnumNameXlCase(xlcApplyStyle);
		EnumNameXlCase(xlcAssignToObject);
		EnumNameXlCase(xlcObjectProtection);
		EnumNameXlCase(xlcHideObject);
		EnumNameXlCase(xlcSetExtract);
		EnumNameXlCase(xlcCreatePublisher);
		EnumNameXlCase(xlcSubscribeTo);
		EnumNameXlCase(xlcAttributes);
		EnumNameXlCase(xlcShowToolbar);
		EnumNameXlCase(xlcPrintPreview);
		EnumNameXlCase(xlcEditColor);
		EnumNameXlCase(xlcShowLevels);
		EnumNameXlCase(xlcFormatMain);
		EnumNameXlCase(xlcFormatOverlay);
		EnumNameXlCase(xlcOnRecalc);
		EnumNameXlCase(xlcEditSeries);
		EnumNameXlCase(xlcDefineStyle);
		EnumNameXlCase(xlcLinePrint);
		EnumNameXlCase(xlcEnterData);
		EnumNameXlCase(xlcGalleryRadar);
		EnumNameXlCase(xlcMergeStyles);
		EnumNameXlCase(xlcEditionOptions);
		EnumNameXlCase(xlcPastePicture);
		EnumNameXlCase(xlcPastePictureLink);
		EnumNameXlCase(xlcSpelling);
		EnumNameXlCase(xlcZoom);
		EnumNameXlCase(xlcResume);
		EnumNameXlCase(xlcInsertObject);
		EnumNameXlCase(xlcWindowMinimize);
		EnumNameXlCase(xlcSize);
		EnumNameXlCase(xlcMove);
		EnumNameXlCase(xlcSoundNote);
		EnumNameXlCase(xlcSoundPlay);
		EnumNameXlCase(xlcFormatShape);
		EnumNameXlCase(xlcExtendPolygon);
		EnumNameXlCase(xlcFormatAuto);
		EnumNameXlCase(xlcGallery3dBar);
		EnumNameXlCase(xlcGallery3dSurface);
		EnumNameXlCase(xlcFillAuto);
		EnumNameXlCase(xlcCustomizeToolbar);
		EnumNameXlCase(xlcAddTool);
		EnumNameXlCase(xlcEditObject);
		EnumNameXlCase(xlcOnDoubleclick);
		EnumNameXlCase(xlcOnEntry);
		EnumNameXlCase(xlcWorkbookAdd);
		EnumNameXlCase(xlcWorkbookMove);
		EnumNameXlCase(xlcWorkbookCopy);
		EnumNameXlCase(xlcWorkbookOptions);
		EnumNameXlCase(xlcSaveWorkspace);
		EnumNameXlCase(xlcChartWizard);
		EnumNameXlCase(xlcDeleteTool);
		EnumNameXlCase(xlcMoveTool);
		EnumNameXlCase(xlcWorkbookSelect);
		EnumNameXlCase(xlcWorkbookActivate);
		EnumNameXlCase(xlcAssignToTool);
		EnumNameXlCase(xlcCopyTool);
		EnumNameXlCase(xlcResetTool);
		EnumNameXlCase(xlcConstrainNumeric);
		EnumNameXlCase(xlcPasteTool);
		EnumNameXlCase(xlcPlacement);
		EnumNameXlCase(xlcFillWorkgroup);
		EnumNameXlCase(xlcWorkbookNew);
		EnumNameXlCase(xlcScenarioCells);
		EnumNameXlCase(xlcScenarioDelete);
		EnumNameXlCase(xlcScenarioAdd);
		EnumNameXlCase(xlcScenarioEdit);
		EnumNameXlCase(xlcScenarioShow);
		EnumNameXlCase(xlcScenarioShowNext);
		EnumNameXlCase(xlcScenarioSummary);
		EnumNameXlCase(xlcPivotTableWizard);
		EnumNameXlCase(xlcPivotFieldProperties);
		EnumNameXlCase(xlcPivotField);
		EnumNameXlCase(xlcPivotItem);
		EnumNameXlCase(xlcPivotAddFields);
		EnumNameXlCase(xlcOptionsCalculation);
		EnumNameXlCase(xlcOptionsEdit);
		EnumNameXlCase(xlcOptionsView);
		EnumNameXlCase(xlcAddinManager);
		EnumNameXlCase(xlcMenuEditor);
		EnumNameXlCase(xlcAttachToolbars);
		EnumNameXlCase(xlcVbaactivate);
		EnumNameXlCase(xlcOptionsChart);
		EnumNameXlCase(xlcVbaInsertFile);
		EnumNameXlCase(xlcVbaProcedureDefinition);
		EnumNameXlCase(xlcRoutingSlip);
		EnumNameXlCase(xlcRouteDocument);
		EnumNameXlCase(xlcMailLogon);
		EnumNameXlCase(xlcInsertPicture);
		EnumNameXlCase(xlcEditTool);
		EnumNameXlCase(xlcGalleryDoughnut);
		EnumNameXlCase(xlcChartTrend);
		EnumNameXlCase(xlcPivotItemProperties);
		EnumNameXlCase(xlcWorkbookInsert);
		EnumNameXlCase(xlcOptionsTransition);
		EnumNameXlCase(xlcOptionsGeneral);
		EnumNameXlCase(xlcFilterAdvanced);
		EnumNameXlCase(xlcMailAddMailer);
		EnumNameXlCase(xlcMailDeleteMailer);
		EnumNameXlCase(xlcMailReply);
		EnumNameXlCase(xlcMailReplyAll);
		EnumNameXlCase(xlcMailForward);
		EnumNameXlCase(xlcMailNextLetter);
		EnumNameXlCase(xlcDataLabel);
		EnumNameXlCase(xlcInsertTitle);
		EnumNameXlCase(xlcFontProperties);
		EnumNameXlCase(xlcMacroOptions);
		EnumNameXlCase(xlcWorkbookHide);
		EnumNameXlCase(xlcWorkbookUnhide);
		EnumNameXlCase(xlcWorkbookDelete);
		EnumNameXlCase(xlcWorkbookName);
		EnumNameXlCase(xlcGalleryCustom);
		EnumNameXlCase(xlcAddChartAutoformat);
		EnumNameXlCase(xlcDeleteChartAutoformat);
		EnumNameXlCase(xlcChartAddData);
		EnumNameXlCase(xlcAutoOutline);
		EnumNameXlCase(xlcTabOrder);
		EnumNameXlCase(xlcShowDialog);
		EnumNameXlCase(xlcSelectAll);
		EnumNameXlCase(xlcUngroupSheets);
		EnumNameXlCase(xlcSubtotalCreate);
		EnumNameXlCase(xlcSubtotalRemove);
		EnumNameXlCase(xlcRenameObject);
		EnumNameXlCase(xlcWorkbookScroll);
		EnumNameXlCase(xlcWorkbookNext);
		EnumNameXlCase(xlcWorkbookPrev);
		EnumNameXlCase(xlcWorkbookTabSplit);
		EnumNameXlCase(xlcFullScreen);
		EnumNameXlCase(xlcWorkbookProtect);
		EnumNameXlCase(xlcScrollbarProperties);
		EnumNameXlCase(xlcPivotShowPages);
		EnumNameXlCase(xlcTextToColumns);
		EnumNameXlCase(xlcFormatCharttype);
		EnumNameXlCase(xlcLinkFormat);
		EnumNameXlCase(xlcTracerDisplay);
		EnumNameXlCase(xlcTracerNavigate);
		EnumNameXlCase(xlcTracerClear);
		EnumNameXlCase(xlcTracerError);
		EnumNameXlCase(xlcPivotFieldGroup);
		EnumNameXlCase(xlcPivotFieldUngroup);
		EnumNameXlCase(xlcCheckboxProperties);
		EnumNameXlCase(xlcLabelProperties);
		EnumNameXlCase(xlcListboxProperties);
		EnumNameXlCase(xlcEditboxProperties);
		EnumNameXlCase(xlcPivotRefresh);
		EnumNameXlCase(xlcLinkCombo);
		EnumNameXlCase(xlcOpenText);
		EnumNameXlCase(xlcHideDialog);
		EnumNameXlCase(xlcSetDialogFocus);
		EnumNameXlCase(xlcEnableObject);
		EnumNameXlCase(xlcPushbuttonProperties);
		EnumNameXlCase(xlcSetDialogDefault);
		EnumNameXlCase(xlcFilter);
		EnumNameXlCase(xlcFilterShowAll);
		EnumNameXlCase(xlcClearOutline);
		EnumNameXlCase(xlcFunctionWizard);
		EnumNameXlCase(xlcAddListItem);
		EnumNameXlCase(xlcSetListItem);
		EnumNameXlCase(xlcRemoveListItem);
		EnumNameXlCase(xlcSelectListItem);
		EnumNameXlCase(xlcSetControlValue);
		EnumNameXlCase(xlcSaveCopyAs);
		EnumNameXlCase(xlcOptionsListsAdd);
		EnumNameXlCase(xlcOptionsListsDelete);
		EnumNameXlCase(xlcSeriesAxes);
		EnumNameXlCase(xlcSeriesX);
		EnumNameXlCase(xlcSeriesY);
		EnumNameXlCase(xlcErrorbarX);
		EnumNameXlCase(xlcErrorbarY);
		EnumNameXlCase(xlcFormatChart);
		EnumNameXlCase(xlcSeriesOrder);
		EnumNameXlCase(xlcMailLogoff);
		EnumNameXlCase(xlcClearRoutingSlip);
		EnumNameXlCase(xlcAppActivateMicrosoft);
		EnumNameXlCase(xlcMailEditMailer);
		EnumNameXlCase(xlcOnSheet);
		EnumNameXlCase(xlcStandardWidth);
		EnumNameXlCase(xlcScenarioMerge);
		EnumNameXlCase(xlcSummaryInfo);
		EnumNameXlCase(xlcFindFile);
		EnumNameXlCase(xlcActiveCellFont);
		EnumNameXlCase(xlcEnableTipwizard);
		EnumNameXlCase(xlcVbaMakeAddin);
		EnumNameXlCase(xlcInsertdatatable);
		EnumNameXlCase(xlcWorkgroupOptions);
		EnumNameXlCase(xlcMailSendMailer);
		EnumNameXlCase(xlcAutocorrect);
		EnumNameXlCase(xlcPostDocument);
		EnumNameXlCase(xlcPicklist);
		EnumNameXlCase(xlcViewShow);
		EnumNameXlCase(xlcViewDefine);
		EnumNameXlCase(xlcViewDelete);
		EnumNameXlCase(xlcSheetBackground);
		EnumNameXlCase(xlcInsertMapObject);
		EnumNameXlCase(xlcOptionsMenono);
		EnumNameXlCase(xlcNormal);
		EnumNameXlCase(xlcLayout);
		EnumNameXlCase(xlcRmPrintArea);
		EnumNameXlCase(xlcClearPrintArea);
		EnumNameXlCase(xlcAddPrintArea);
		EnumNameXlCase(xlcMoveBrk);
		EnumNameXlCase(xlcHidecurrNote);
		EnumNameXlCase(xlcHideallNotes);
		EnumNameXlCase(xlcDeleteNote);
		EnumNameXlCase(xlcTraverseNotes);
		EnumNameXlCase(xlcActivateNotes);
		EnumNameXlCase(xlcProtectRevisions);
		EnumNameXlCase(xlcUnprotectRevisions);
		EnumNameXlCase(xlcOptionsMe);
		EnumNameXlCase(xlcWebPublish);
		EnumNameXlCase(xlcNewwebquery);
		EnumNameXlCase(xlcPivotTableChart);
		EnumNameXlCase(xlcOptionsSave);
		EnumNameXlCase(xlcOptionsSpell);
		EnumNameXlCase(xlcHideallInkannots);
#endif
		default:
			stream << nFunc;
			break;
	}	
	str = stream.str();
}

void LogHelper::GetXlFunctionTypeStr(int xlfn, std::wstring& str)
{
	XLCALL_FUNCTYPE type = (XLCALL_FUNCTYPE)(FuncTypeMask & xlfn);
	str.clear();
	if (type & xlPrompt)
		str += __X("[Prompt]");
	if (type & xlIntl)
		str += __X("[Intl]");
	if (type & xlSpecial)
		str += __X("[Special]");
	if (type & xlCommand)
		str += __X("[Command]");
}

void LogHelper::GetXlResultName(XLCALL_RESULT res, std::wstring& str)
{
	std::wstringstream stream;
	switch (res)
	{
		EnumNameCase2(xlret, Success);
		EnumNameCase2(xlret, Abort);
		EnumNameCase2(xlret, InvXlfn);
		EnumNameCase2(xlret, InvCount);
		EnumNameCase2(xlret, InvXloper);
		EnumNameCase2(xlret, StackOvfl);
		EnumNameCase2(xlret, Failed);
		EnumNameCase2(xlret, Uncalced);
		EnumNameCase2(xlret, NotThreadSafe);
		EnumNameCase2(xlret, InvAsynchronousContext);
		EnumNameCase2(xlret, NotClusterSafe);
	default:
		stream << res;
		break;
	}
	str = stream.str();
}

void LogHelper::GetXloperTypeName(int type, std::wstring& str)
{
	std::wstringstream stream;

	if (type & xlbitDLLFree)
		stream << __X("[DLLFree]");
	if (type & xlbitXLFree)
		stream << __X("[XLFree]");

	type = (XLOPERTYPE)(XLOPER_TYPEMASK & type);
	switch (type)
	{
		EnumNameCase2(xltype, Num);
		EnumNameCase2(xltype, Str);
		EnumNameCase2(xltype, Bool);
		EnumNameCase2(xltype, Ref);
		EnumNameCase2(xltype, Err);
		EnumNameCase2(xltype, Flow);
		EnumNameCase2(xltype, Multi);
		EnumNameCase2(xltype, Missing);
		EnumNameCase2(xltype, Nil);
		EnumNameCase2(xltype, SRef);
		EnumNameCase2(xltype, Int);
		EnumNameCase2(xltype, BigData);
	default:
		stream << type;
		break;
	}
	
	str = stream.str();
}

void LogHelper::GetXloperErrName(XLOPER_ERRTYPE type, std::wstring& str)
{
	std::wstringstream stream;
	switch (type)
	{
		EnumNameCase(xlerrNull);
		EnumNameCase(xlerrDiv0);
		EnumNameCase(xlerrValue);
		EnumNameCase(xlerrRef);
		EnumNameCase(xlerrName);
		EnumNameCase(xlerrNum);
		EnumNameCase(xlerrNA);
		EnumNameCase(xlerrGettingData);
	default:
		stream << type;
		break;
	}
	str = stream.str();
}

void LogHelper::GetPascalString(LPCSTR pStr, std::wstring& result)
{
	result.clear();
	if (pStr)
	{
		UINT nLen = (BYTE)pStr[0];
		UINT nNewLen = MultiByteToWideChar(CP_ACP, 0, pStr + 1, nLen, NULL, NULL);
		WCHAR *pNewStr = (WCHAR*)malloc((nNewLen + 1) * sizeof(WCHAR));
		if (pNewStr)
		{
			MultiByteToWideChar(CP_ACP, 0, pStr + 1, nLen, pNewStr, nNewLen);
			pNewStr[nNewLen] = __Xc('\0');
			result = pNewStr;
		}
		free(pNewStr);
	}
}

void LogHelper::StrToWStr(LPCSTR pStr, std::wstring& result)
{
	result.clear();
	if (pStr)
	{
		UINT nLen = strlen(pStr);
		UINT nNewLen = MultiByteToWideChar(CP_ACP, 0, pStr, nLen, NULL, NULL);
		WCHAR *pNewStr = (WCHAR*)malloc((nNewLen + 1) * sizeof(WCHAR));
		if (pNewStr)
		{
			MultiByteToWideChar(CP_ACP, 0, pStr, nLen, pNewStr, nNewLen);
			pNewStr[nNewLen] = __Xc('\0');
			result = pNewStr;
		}
		free(pNewStr);
	}
}

BOOL LogHelper::WStrToStr(const std::wstring& wstr, std::string& str)
{
	str.clear();
	UINT nLen = wstr.size();
	UINT nNewLen = WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), nLen, NULL, 0, NULL, NULL);
	if (nNewLen > PascalStrMaxLen)
		nNewLen = PascalStrMaxLen;

	LPSTR pNewStr = (LPSTR)malloc((nNewLen + 1) * sizeof(char));
	if (pNewStr)
	{
		WideCharToMultiByte(CP_ACP, 0, wstr.c_str(), nLen, pNewStr, nNewLen, NULL, NULL);
		pNewStr[nNewLen] = '\0';
		str = pNewStr;
		free(pNewStr);
		return TRUE;
	}
	return FALSE;
}

void LogHelper::GetPascalString(LPCWSTR lpStr, std::wstring& result)
{
	result.clear();
	if (lpStr)
	{
		UINT nSize = *lpStr;
		if (nSize > 0)
		{
			result.assign(lpStr, 1, nSize);
		}
	}
}

void LogHelper::OpenLogFile()
{
	if (m_fileStream.is_open())
		m_fileStream.close();

	WCHAR pwszPath[MAX_PATH] = { 0 };
	::SHGetSpecialFolderPathW(NULL, pwszPath, CSIDL_MYDOCUMENTS, TRUE);
	m_sLogPath = pwszPath;

	std::wstring sTime;
	if (sTime.empty())
	{
		time_t rawtime;
		struct tm timeinfo;

		time(&rawtime);
		localtime_s(&timeinfo, &rawtime);
		std::wstringstream stream;
		stream << std::put_time(&timeinfo, __X("%Y%m%d_%H%M%S"));
		sTime = stream.str();
	}

	m_sLogPath += __X("\\xllhook\\");
	::CreateDirectoryW(m_sLogPath.c_str(), NULL);

	m_sLogPath += sTime.c_str();
	::CreateDirectoryW(m_sLogPath.c_str(), NULL);

	m_sLogFile = m_sLogPath + __X("\\api.htm");
	m_fileStream.open(m_sLogFile.c_str(), std::wfstream::out | std::wfstream::app);
	m_fileStream.imbue(std::locale(""));

	m_fileStream << TableBegin << std::endl;
	PrintLogTitle();
}

void LogHelper::CloseLogFile()
{
	if (m_fileStream.is_open())
	{
		m_fileStream << TableEnd;
		m_fileStream.close();
		m_bFirstLog = true;
	}
}

void LogHelper::ClearLog()
{
	if (!m_fileStream.is_open())
		return;
	m_fileStream.close();
	m_fileStream.open(m_sLogFile.c_str(), std::wfstream::out);
	m_fileStream.imbue(std::locale(""));
	m_fileStream << TableBegin << std::endl;
	PrintLogTitle();
}

void LogHelper::OpenFolder()
{
	HINSTANCE res = ::ShellExecuteW(0, __X("open"), m_sLogPath.c_str(), NULL, NULL, SW_SHOW);
	DWORD err = GetLastError();
}

void LogHelper::PrintLogTitle()
{
	if (!m_fileStream.is_open())
		return;

	m_fileStream << RowBegin
		<< ColBegin << __X("FuncAttr") << ColEnd
		<< ColBegin << __X("FuncName") << ColEnd
		<< ColBegin << __X("FuncRes") << ColEnd
		<< ColBegin << __X("ResType") << ColEnd
		<< ColBegin << __X("ResValue") << ColEnd;
	for (UINT i = 1; i <= 30; ++i)
	{
		m_fileStream
			<< ColBegin << __X("Type") << i << ColEnd
			<< ColBegin << __X("Value") << i << ColEnd;
	}
	m_fileStream << RowEnd << std::endl;
}

void LogHelper::PrintTopBuffer(UINT deep)
{
	if (m_bFirstLog)
	{
		OpenLogFile();
		m_bFirstLog = false;
	}

	++m_nLineCount;
	if (!m_fileStream.is_open())
		return;

	if (m_fileStream.bad())
		m_fileStream.clear();

	LogBuffer buffer;
	std::swap(buffer, m_callstack.back());
	m_callstack.pop_back();

	if (!m_callstack.empty() && m_callstack.back().bPrintEnter)
	{
		m_callstack.back().bPrintEnter = false;
		PrintEnterRow(m_callstack.back().sFuncName);
		PrintTopBuffer(deep + 1);
	}
	if (deep > 0)
		m_callstack.push_back(buffer);

	m_logFileMutex.lock();
	m_fileStream << RowBegin
		<< ColBegin << buffer.sFuncAttr << ColEnd
		<< ColBegin << buffer.sFuncName << ColEnd
		<< ColBegin << buffer.sResult << ColEnd
		<< ColBegin << buffer.sResOperType << ColEnd
		<< ColBegin << buffer.sResOperValue << ColEnd;
	ASSERT(!m_fileStream.bad());

	UINT nArgNum = min(buffer.argsOperType.size(), buffer.argsOperValue.size());
	for (UINT i = 0; i < nArgNum; ++i)
	{
		m_fileStream
			<< ColBegin << buffer.argsOperType[i] << ColEnd
			<< ColBegin << buffer.argsOperValue[i] << ColEnd;
		ASSERT(!m_fileStream.bad());
	}

	m_fileStream << RowEnd << std::endl;
	ASSERT(!m_fileStream.bad());
	if (m_fileStream.bad())
		m_fileStream.clear();

	m_logFileMutex.unlock();

	if (!buffer.bPrintEnter && 0 == deep)
		PrintLeaveRow(buffer.sFuncName);
// 	buffer.clear();
}

void LogHelper::PrintEnterRow(const std::wstring& name)
{
	m_logFileMutex.lock();
	m_fileStream << RowBegin
		<< ColBegin << __X("#Enter#") << ColEnd
		<< ColBegin << __X("#Enter#") << ColEnd
		<< ColBegin << name << ColEnd
		<< ColBegin << __X("#Enter#") << ColEnd
		<< RowEnd << std::endl;
	m_logFileMutex.unlock();
}

void LogHelper::PrintLeaveRow(const std::wstring& name)
{
	m_logFileMutex.lock();
	m_fileStream << RowBegin
		<< ColBegin << __X("#Leave#") << ColEnd
		<< ColBegin << __X("#Leave#") << ColEnd
		<< ColBegin << name << ColEnd
		<< ColBegin << __X("#Leave#") << ColEnd
		<< RowEnd << std::endl;
	m_logFileMutex.unlock();
}

void LogHelper::LogLPenHelperBegin(int wCode, void* lpv)
{
	if (!IsNeedLog())
		return;

	m_callstack.push_back(LogBuffer());
	LogBuffer& buffer = m_callstack.back();
	buffer.xlfn = wCode;

	buffer.sFuncName = __X("LPenHelper");
	buffer.argsOperType.resize(2);
	buffer.argsOperValue.resize(2);

	std::wstringstream stream;
	stream << wCode;
	buffer.argsOperValue[0] = stream.str();

	stream.str(std::wstring());
	stream << __X("0x") << std::hex << lpv << std::dec;
	buffer.argsOperValue[1] = stream.str();
}

void LogHelper::LogLPenHelperEnd(int result)
{
	if (!IsNeedLog())
		return;

	LogBuffer& buffer = m_callstack.back();

	std::wstringstream stream;
	stream << result;
	buffer.sResult = stream.str();
	PrintTopBuffer(0);
}

void LogHelper::RegisterFunction(LogBuffer& buffer)
{
#if SET_Hook_XLLExport
	if (m_nCodePos >= nMaxUDFuncNum)
#endif
		return;

	if (buffer.argsOperValue.size() < 3)
		return;

	const std::wstring& sModule = buffer.argsOperValue[0];
	const std::wstring& sProcedure = buffer.argsOperValue[1];
	const std::wstring& sTypeText = buffer.argsOperValue[2];

	std::string sProc;
	WStrToStr(sProcedure, sProc);
	if (sProc.empty())
		return;

	HMODULE hModule = ::GetModuleHandleW(sModule.c_str());
	void* lpProc = ::GetProcAddress(hModule, sProc.c_str());
	if (!lpProc && !sProc.empty())
	{
		WORD nExportID = atoi(sProc.c_str());
		void* lpProc = ::GetProcAddress(hModule, (LPCSTR)nExportID);
	}
	if (!lpProc)
		return;

	if (m_udfMap.find(lpProc) != m_udfMap.end())
		return;

	XllFuncInfo& info = m_udfMap[lpProc];
	HRESULT hr = ParseArgumentType(sTypeText.c_str(), info);
	if (FAILED(hr))
		return;

// 	AttachFunction(&lpProc, (PVOID)UDFHook, &sProc[0]);

	m_codes[m_nCodePos] = ShellCode(lpProc);
	AttachFunction(&lpProc, m_codes[m_nCodePos].address(), &sProc[0]);
	++m_nCodePos;
	info.pEntryPoint = lpProc;
	info.funcName = sProcedure;
}

#define RESOLUT_TYPE(xl11type, xl12type)\
	if (__Xc('%') == *(lpArgType + 1))	\
	{									\
		++lpArgType;					\
		curType = (xl12type);			\
	}									\
	else								\
	{									\
		curType = (xl11type);			\
	}

HRESULT LogHelper::ParseArgumentType(LPCWSTR lpArgType, XllFuncInfo& info)
{
	if (!lpArgType || 0 == lpArgType[0])
		return E_INVALIDARG;

	bool bReturnType = true;
	while (lpArgType && __Xc('\0') != *lpArgType)
	{
		XlCallArgType curType = xlArgNone;
		switch (*lpArgType)
		{
		case __Xc('A'): curType = xlArgBool; break;
		case __Xc('L'): curType = xlArgBoolRef; break;
		case __Xc('B'): curType = xlArgDouble; break;
		case __Xc('E'): curType = xlArgDoubleRef; break;
		case __Xc('C'):
			RESOLUT_TYPE(xlArgCStr, xlArgCWStr);
			break;
		case __Xc('F'):
			RESOLUT_TYPE(xlArgCStrInOut, xlArgCWStrInOut);
			break;
		case __Xc('D'):
			RESOLUT_TYPE(xlArgPascalStr, xlArgPascalWStr);
			break;
		case __Xc('G'):
			RESOLUT_TYPE(xlArgPascalStrInOut, xlArgPascalWStrInOut);
			break;
		case __Xc('H'): curType = xlArgUShort; break;
		case __Xc('I'): curType = xlArgShort; break;
		case __Xc('M'): curType = xlArgShortRef; break;
		case __Xc('J'): curType = xlArgInt; break;
		case __Xc('N'): curType = xlArgIntRef; break;
		case __Xc('K'):
			RESOLUT_TYPE(xlArgFloatArr, xlArgFloatArr12);
			break;
		case __Xc('O'):
			RESOLUT_TYPE(xlArgArray, xlArgArray12);
			break;
		case __Xc('P'): curType = xlArgOper; break;
		case __Xc('R'): curType = xlArgXLOper; break;
		case __Xc('Q'): curType = xlArgOper12; break;
		case __Xc('U'): curType = xlArgXLOper12; break;
		case __Xc('!'): info.funcAttr |= xlArgBitVolatile; break;
		case __Xc('#'): info.funcAttr |= xlArgBitMacroFunc; break;
		case __Xc('$'): info.funcAttr |= xlArgBitThreadSafe; break;
		case __Xc('&'): info.funcAttr |= xlArgBitClusterSafe; break;
		default:
			if (bReturnType)
			{
				curType = ParseVoidRet(*lpArgType);
				if (xlArgNone != curType)
					break;
			}
			return E_FAIL;
		}

		if (!bReturnType)
		{
			if (xlArgNone != curType)
				info.argTypes.push_back(curType);
		}
		else
		{
			// 返回值不允许使用有3个参数的O类型
			if (xlArgArray == curType || xlArgArray12 == curType)
				return E_INVALIDARG;

			info.retrunType = curType;
			bReturnType = false;
		}
		++lpArgType;
	}
	return S_OK;
}

XlCallArgType LogHelper::ParseVoidRet(WCHAR typeChar)
{
	if (__Xc('<') == typeChar)
	{
		return xlArgRetrun1;
	}
	else if (__Xc('1') <= typeChar &&
		__Xc('9') >= typeChar)
	{
		return (XlCallArgType)(typeChar - __Xc('0'));
	}
	return xlArgNone;
}

void** LogHelper::LogUdfArgument(void* key, void** lpArgument)
{
	m_callstack.push_back(LogBuffer());
	LogBuffer& buff = m_callstack.back();

	const XllFuncInfo& info = m_udfMap.at(key);
	const std::vector<XlCallArgType>& types = info.argTypes;
	buff.sFuncAttr = __X("Excel");
	buff.sFuncName = info.funcName;
	buff.argsOperType.resize(types.size());
	buff.argsOperValue.resize(types.size());
	for (UINT i = 0; i < types.size(); ++i)
	{
		GetUDFArgTypeName(types[i], buff.argsOperType[i]);
		DWORD addrInc = GetUDFArgValue(types[i], lpArgument, buff.argsOperValue[i]);
		lpArgument += addrInc;
	}
	return lpArgument;
}

void LogHelper::GetUDFArgTypeName(XlCallArgType type, std::wstring& name)
{
	name.clear();
	switch (type)
	{
	case xlArgNone:
		name = __X("void");
		break;
	case xlArgRetrun1:
		name = __X("arg1");
		break;
	case xlArgRetrun2:
		name = __X("arg2");
		break;
	case xlArgRetrun3:
		name = __X("arg3");
		break;
	case xlArgRetrun4:
		name = __X("arg4");
		break;
	case xlArgRetrun5:
		name = __X("arg5");
		break;
	case xlArgRetrun6:
		name = __X("arg6");
		break;
	case xlArgRetrun7:
		name = __X("arg7");
		break;
	case xlArgRetrun8:
		name = __X("arg8");
		break;
	case xlArgRetrun9:
		name = __X("arg9");
		break;
	case xlArgBool:
		name = __X("BOOL");
		break;
	case xlArgBoolRef:
		name = __X("BOOL*");
		break;
	case xlArgDouble:
		name = __X("double");
		break;
	case xlArgDoubleRef:
		name = __X("double*");
		break;
	case xlArgCStr:
		name = __X("char*");
		break;
	case xlArgPascalStr:
		name = __X("char*[pas]");
		break;
	case xlArgUShort:
		name = __X("ushort");
		break;
	case xlArgShort:
		name = __X("short");
		break;
	case xlArgShortRef:
		name = __X("short*");
		break;
	case xlArgInt:
		name = __X("int");
		break;
	case xlArgIntRef:
		name = __X("int*");
		break;
	case xlArgFloatArr:
		name = __X("FP*");
		break;
	case xlArgArray:
		name = __X("ushort*, ushort*, double*");
		break;
	case xlArgOper:
		name = __X("oper*");
		break;
	case xlArgXLOper:
		name = __X("xloper*");
		break;
	case xlArgCWStr:
		name = __X("WCHAR*");
		break;
	case xlArgPascalWStr:
		name = __X("WCHAR*[pas]");
		break;
	case xlArgFloatArr12:
		name = __X("FP12*");
		break;
	case xlArgArray12:
		name = __X("int*, int*, double*");
		break;
	case xlArgOper12:
		name = __X("oper12*");
		break;
	case xlArgXLOper12:
		name = __X("xloper12*");
		break;
	case xlArgCStrInOut:
		name = __X("char*[modify]");
		break;
	case xlArgCWStrInOut:
		name = __X("WCHAR*[modify]");
		break;
	case xlArgPascalStrInOut:
		name = __X("char*[pas][modify]");
		break;
	case xlArgPascalWStrInOut:
		name = __X("WCHAR*[pas][modify]");
		break;
	default:
		break;
	}
}

DWORD LogHelper::GetUDFArgValue(XlCallArgType type, void** lpArgument, std::wstring& value)
{
	DWORD addressInc = 1;
	std::wstringstream stream;
	switch (type)
	{
	case xlArgBool:
		stream << (*((BOOL*)lpArgument) ? __X("TRUE") : __X("FALSE"));
		break;
	case xlArgBoolRef:
		stream << (**((BOOL**)lpArgument) ? __X("TRUE") : __X("FALSE"));
		break;
	case xlArgDouble:
		stream << *((double*)lpArgument);
		++addressInc;
		break;
	case xlArgDoubleRef:
		stream << **((double**)lpArgument);
		break;
	case xlArgCStr:
	case xlArgCStrInOut:
		stream << *((char**)lpArgument);
		break;
	case xlArgPascalStr:
	case xlArgPascalStrInOut:
	{
		std::wstring str;
		GetPascalString(*((char**)lpArgument), str);
		stream << str;
		break;
	}
	case xlArgUShort:
		stream << *((unsigned short*)lpArgument);
		break;
	case xlArgShort:
		stream << *((short*)lpArgument);
		break;
	case xlArgShortRef:
		stream << **((short**)lpArgument);
		break;
	case xlArgInt:
		stream << *((int*)lpArgument);
		break;
	case xlArgIntRef:
		stream << **((int**)lpArgument);
		break;
	case xlArgFloatArr:
	{
		FP* pFP = *((FP**)lpArgument);
		stream << pFP->rows << __Xc('x') << pFP->columns;
		break;
	}
	case xlArgArray:
	{
		WORD rows = **((WORD**)lpArgument);
		++lpArgument;
		WORD cols = **((WORD**)lpArgument);
		++lpArgument;
		double* lparray = *((double**)lpArgument);
		stream << rows << __Xc('x') << cols;
		addressInc += 2;
		break;
	}
	case xlArgOper:
	case xlArgXLOper:
	{
		LPXLOPER lpoper = *((LPXLOPER*)lpArgument);
		std::wstring sType, sVal;
		LogXloper(lpoper, sType, sVal);
		stream << __Xc('[') << sType << __X("] ") << sVal;
		break;
	}
	case xlArgCWStr:
	case xlArgCWStrInOut:
		stream << *((WCHAR**)lpArgument);
		break;
	case xlArgPascalWStr:
	case xlArgPascalWStrInOut:
	{
		std::wstring str;
		GetPascalString(*((WCHAR**)lpArgument), str);
		stream << str;
		break;
	}
	case xlArgFloatArr12:
	{
		FP12* pFP = *((FP12**)lpArgument);
		stream << pFP->rows << __Xc('x') << pFP->columns;
		break;
	}
	case xlArgArray12:
	{
		int rows = **((int**)lpArgument);
		++lpArgument;
		int columns = **((int**)lpArgument);
		++lpArgument;
		double* lparray = *((double**)lpArgument);
		stream << rows << __Xc('x') << columns;
		addressInc += 2;
		break;
	}
	case xlArgOper12:
	case xlArgXLOper12:
	{
		LPXLOPER12 lpoper = *((LPXLOPER12*)lpArgument);
		std::wstring sType, sVal;
		LogXloper(lpoper, sType, sVal);
		stream << __Xc('[') << sType << __X("] ") << sVal;
		break;
	}
	default:
		break;
	}
	value = stream.str();
	return addressInc;
}

void LogHelper::LogUdfEnd(void* key, XlFuncResult result)
{
	if (!IsNeedLog())
	{
		m_callstack.pop_back();
		return;
	}

	LogBuffer& buff = m_callstack.back();
	const XllFuncInfo& info = m_udfMap.at(key);
	GetUDFArgTypeName(info.retrunType, buff.sResOperType);
	GetUDFArgValue(info.retrunType, (void**)&result, buff.sResOperValue);
	PrintTopBuffer(0);
}

DWORD PASCAL UDFHook()
{
	PVOID* lpArgBegin = NULL;
	__asm
	{
		mov eax, ebp;
		add	eax, 8;
		mov lpArgBegin, eax;
	}
	XlFuncResult funcRes = { 0 };
	XlCallArgType resType = xlArgNone;
	UINT nBytes = 0;
	{
		PVOID key = *lpArgBegin;
		++lpArgBegin;

		PVOID* lpArgEnd = LogHelper::Instance().LogUdfArgument(key, lpArgBegin);
		const XllFuncInfo& info = LogHelper::Instance().GetUDFMap().at(key);

		resType = info.retrunType;
		void* pEntry = info.pEntryPoint;
		UINT nDWORDs = lpArgEnd - lpArgBegin;
		nBytes = nDWORDs * 4;
		DWORD nESP = 0;
		__asm
		{
			mov		eax, lpArgBegin;
			mov		ecx, nDWORDs;
			mov		nESP, esp;
			sub		esp, nBytes;
			mov		edx, esp;
			push	ebx;
			{
			LoopBegin:
				cmp		ecx, 0;
				je		LoopEnd;
				mov		ebx, [eax];
				mov		[edx], ebx;

				dec		ecx;
				add		eax, 4;
				add		edx, 4;
				jmp		LoopBegin;
			LoopEnd:
			}
			pop		ebx;
			call	pEntry;
			mov		esp, nESP;
			{
				cmp		resType, xlArgDouble;
				je		GetDouble;
				// default:
				mov		dword ptr[funcRes], eax;
				jmp		EndCall;
			GetDouble:	// case xlArgDouble
				fstp	qword ptr[funcRes];
			}
		EndCall:
		}
		LogHelper::Instance().LogUdfEnd(key, funcRes);
	}

	__asm
	{
		cmp		resType, xlArgDouble;
		je		ReturnDouble;
		// default:
		mov		eax, dword ptr[funcRes];
		jmp		EndReturn;
	ReturnDouble:	// case xlArgDouble
		fld		qword ptr[funcRes];
	EndReturn:
		push	ebx;
		mov		ecx, lpArgBegin;
		add		ecx, nBytes;
		mov		edx, ebp;
		add		edx, 8;
		{
		LoopBegin2:
			cmp		esp, edx;
			je		LoopEnd2;
			sub		ecx, 4;
			sub		edx, 4;
			mov		ebx, [edx];
			mov		[ecx], ebx;
			jmp		LoopBegin2;
		LoopEnd2:
		}
		sub		ecx, edx;
		add		ebp, ecx;
		add		esp, ecx;
		pop		ebx;
// 		mov		ecx, __security_cookie;
// 		xor		ecx, ebp;
// 		mov		[ebp - 10h], ecx;
	}
	return funcRes.dw;
}

ShellCode::ShellCode(PVOID srcFunc)
{
	char* lpAddr = m_code;

	(*(BYTE*)lpAddr) = 0xB8;	// mov eax, srcFunc
	++lpAddr;
	*((PVOID*)lpAddr) = srcFunc;
	lpAddr += sizeof(PVOID);

	(*(BYTE*)lpAddr) = (char)0x59;	// pop ecx
	++lpAddr;
	(*(BYTE*)lpAddr) = (char)0x50;	// push eax
	++lpAddr;
	(*(BYTE*)lpAddr) = (char)0x51;	// push ecx
	++lpAddr;

	(*(BYTE*)lpAddr) = 0xB8;	// mov eax, UDFHook
	++lpAddr;
	*((PVOID*)lpAddr) = UDFHook;
	lpAddr += sizeof(PVOID);

	(*(BYTE*)lpAddr) = (char)0xFF;	// jmp, eax
	++lpAddr;
	(*(BYTE*)lpAddr) = (char)0xE0;
	++lpAddr;

	ASSERT(lpAddr - m_code < m_size);
}

PVOID ShellCode::address()
{
	return (PVOID)m_code;
}

void ShellCode::operator=(const ShellCode& other)
{
	memcpy(m_code, other.m_code, m_size);
}
