#include "loghelper.h"
#include <ShlObj.h>
#include <ctime>

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
	, m_bPause(false)
{
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
		UINT nLen = pStr[0];
		WCHAR *pNewStr = (WCHAR*)malloc((nLen + 1) * sizeof(WCHAR));
		if (pNewStr && nLen > 0)
		{
			MultiByteToWideChar(CP_ACP, 0, &pStr[1], nLen, pNewStr, nLen);
			pNewStr[nLen] = __Xc('\0');
			result = pNewStr;
		}
		free(pNewStr);
	}
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

	time_t rawtime;
	struct tm timeinfo;

	time(&rawtime);
	localtime_s(&timeinfo, &rawtime);
	std::wstringstream stream;
	stream << std::put_time(&timeinfo, __X("%Y%m%d_%H%M%S"));

	m_sLogPath += __X("\\xllhook\\");
	::CreateDirectoryW(m_sLogPath.c_str(), NULL);

	m_sLogPath += stream.str();
	::CreateDirectoryW(m_sLogPath.c_str(), NULL);

	m_sLogFile = m_sLogPath + __X("\\api.csv");
	m_fileStream.open(m_sLogFile.c_str(), std::wfstream::out);
	m_fileStream.imbue(std::locale(""));
	PrintLogTitle();
}

void LogHelper::CloseLogFile()
{
	if (m_fileStream.is_open())
		m_fileStream.close();
}

void LogHelper::PrintLogTitle()
{
	if (!m_fileStream.is_open())
		return;

	static const WCHAR sPreFix[]
		= __X("FuncAttr,FuncName,FuncRes,ResType,ResValue");
	m_fileStream << sPreFix;

	for (UINT i = 1; i <= 30; ++i)
	{
		m_fileStream << __X(",Type") << i
			<< __X(",Value") << i;
	}
	m_fileStream << std::endl;
}

void LogHelper::PrintBuffer(LogBuffer& buffer)
{
	++m_nLineCount;
	if (!m_fileStream.is_open())
		return;

	m_logFileMutex.lock();
	const WCHAR chSep = __Xc(',');
	m_fileStream << buffer.sFuncAttr
		<< chSep << buffer.sFuncName
		<< chSep << buffer.sResult
		<< chSep << buffer.sResOperType
		<< chSep << buffer.sResOperValue;

	UINT nArgNum = min(buffer.argsOperType.size(), buffer.argsOperValue.size());
	for (UINT i = 0; i < nArgNum; ++i)
	{
		m_fileStream
			<< chSep << buffer.argsOperType[i]
			<< chSep << buffer.argsOperValue[i];
	}

	m_fileStream << std::endl;
	m_logFileMutex.unlock();

// 	buffer.clear();
}

void LogHelper::LogLPenHelperBegin(int wCode, void* lpv, LogBuffer& buffer)
{
	if (m_bPause)
		return;

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

void LogHelper::LogLPenHelperEnd(int result, LogBuffer& buffer)
{
	if (m_bPause)
		return;

	std::wstringstream stream;
	stream << result;
	buffer.sResult = stream.str();
}
