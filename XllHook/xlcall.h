/*
**  Microsoft Excel Developer's Toolkit
**  Version 15.0
**
**  File:           INCLUDE\XLCALL.H
**  Description:    Header file for for Excel callbacks
**  Platform:       Microsoft Windows
**
**  DEPENDENCY:
**  Include <windows.h> before you include this.
**
**  This file defines the constants and
**  data types which are used in the
**  Microsoft Excel C API.
**
*/
#ifndef __XLCALL_H__
#define __XLCALL_H__

/*
** XL 12 Basic Datatypes 
**/

typedef INT32 BOOL;			/* Boolean */
typedef WCHAR XCHAR;			/* Wide Character */
typedef INT32 RW;			/* XL 12 Row */
typedef INT32 COL;	 	      	/* XL 12 Column */
typedef DWORD_PTR IDSHEET;		/* XL12 Sheet ID */

/*
** XLREF structure 
**
** Describes a single rectangular reference.
*/

typedef struct xlref 
{
	WORD rwFirst;
	WORD rwLast;
	BYTE colFirst;
	BYTE colLast;
} XLREF, *LPXLREF;


/*
** XLMREF structure
**
** Describes multiple rectangular references.
** This is a variable size structure, default 
** size is 1 reference.
*/

typedef struct xlmref 
{
	WORD count;
	XLREF reftbl[1];					/* actually reftbl[count] */
} XLMREF, *LPXLMREF;


/*
** XLREF12 structure 
**
** Describes a single XL 12 rectangular reference.
*/

typedef struct xlref12
{
	RW rwFirst;
	RW rwLast;
	COL colFirst;
	COL colLast;
} XLREF12, *LPXLREF12;


/*
** XLMREF12 structure
**
** Describes multiple rectangular XL 12 references.
** This is a variable size structure, default 
** size is 1 reference.
*/

typedef struct xlmref12
{
	WORD count;
	XLREF12 reftbl[1];					/* actually reftbl[count] */
} XLMREF12, *LPXLMREF12;


/*
** FP structure
**
** Describes FP structure.
*/

typedef struct _FP
{
    unsigned short int rows;
    unsigned short int columns;
    double array[1];        /* Actually, array[rows][columns] */
} FP;

/*
** FP12 structure
**
** Describes FP structure capable of handling the big grid.
*/

typedef struct _FP12
{
    INT32 rows;
    INT32 columns;
    double array[1];        /* Actually, array[rows][columns] */
} FP12;


/*
** XLOPER structure 
**
** Excel's fundamental data type: can hold data
** of any type. Use "R" as the argument type in the 
** REGISTER function.
**/

typedef struct xloper 
{
	union 
	{
		double num;					/* xltypeNum */
		LPSTR str;					/* xltypeStr */
#ifdef __cplusplus
		WORD xbool;					/* xltypeBool */
#else	
		WORD bool;					/* xltypeBool */
#endif	
		WORD err;					/* xltypeErr */
		short int w;					/* xltypeInt */
		struct 
		{
			WORD count;				/* always = 1 */
			XLREF ref;
		} sref;						/* xltypeSRef */
		struct 
		{
			XLMREF *lpmref;
			IDSHEET idSheet;
		} mref;						/* xltypeRef */
		struct 
		{
			struct xloper *lparray;
			WORD rows;
			WORD columns;
		} array;					/* xltypeMulti */
		struct 
		{
			union
			{
				short int level;		/* xlflowRestart */
				short int tbctrl;		/* xlflowPause */
				IDSHEET idSheet;		/* xlflowGoto */
			} valflow;
			WORD rw;				/* xlflowGoto */
			BYTE col;				/* xlflowGoto */
			BYTE xlflow;
		} flow;						/* xltypeFlow */
		struct
		{
			union
			{
				BYTE *lpbData;			/* data passed to XL */
				HANDLE hdata;			/* data returned from XL */
			} h;
			long cbData;
		} bigdata;					/* xltypeBigData */
	} val;
	WORD xltype;
} XLOPER, *LPXLOPER;

/*
** XLOPER12 structure 
**
** Excel 12's fundamental data type: can hold data
** of any type. Use "U" as the argument type in the 
** REGISTER function.
**/

typedef struct xloper12 
{
	union 
	{
		double num;				       	/* xltypeNum */
		XCHAR *str;				       	/* xltypeStr */
		BOOL xbool;				       	/* xltypeBool */
		int err;				       	/* xltypeErr */
		int w;
		struct 
		{
			WORD count;			       	/* always = 1 */
			XLREF12 ref;
		} sref;						/* xltypeSRef */
		struct 
		{
			XLMREF12 *lpmref;
			IDSHEET idSheet;
		} mref;						/* xltypeRef */
		struct 
		{
			struct xloper12 *lparray;
			RW rows;
			COL columns;
		} array;					/* xltypeMulti */
		struct 
		{
			union
			{
				int level;			/* xlflowRestart */
				int tbctrl;			/* xlflowPause */
				IDSHEET idSheet;		/* xlflowGoto */
			} valflow;
			RW rw;				       	/* xlflowGoto */
			COL col;			       	/* xlflowGoto */
			BYTE xlflow;
		} flow;						/* xltypeFlow */
		struct
		{
			union
			{
				BYTE *lpbData;			/* data passed to XL */
				HANDLE hdata;			/* data returned from XL */
			} h;
			long cbData;
		} bigdata;					/* xltypeBigData */
	} val;
	DWORD xltype;
} XLOPER12, *LPXLOPER12;

/*
** XLOPER and XLOPER12 data types
**
** Used for xltype field of XLOPER and XLOPER12 structures
*/
enum XLOPERTYPE
{
	xltypeInvalid	= 0x0000,
	xltypeNum		= 0x0001,
	xltypeStr		= 0x0002,
	xltypeBool		= 0x0004,
	xltypeRef		= 0x0008,
	xltypeErr		= 0x0010,
	xltypeFlow		= 0x0020,
	xltypeMulti		= 0x0040,
	xltypeMissing	= 0x0080,
	xltypeNil		= 0x0100,
	xltypeSRef		= 0x0400,
	xltypeInt		= 0x0800,

	xlbitXLFree		= 0x1000,
	xlbitDLLFree	= 0x4000,

	xltypeBigData	= (xltypeStr | xltypeInt),
};


/*
** Error codes
**
** Used for val.err field of XLOPER and XLOPER12 structures
** when constructing error XLOPERs and XLOPER12s
*/
enum XLOPER_ERRTYPE
{
	xlerrNull			= 0,
	xlerrDiv0			= 7,
	xlerrValue			= 15,
	xlerrRef			= 23,
	xlerrName			= 29,
	xlerrNum			= 36,
	xlerrNA				= 42,
	xlerrGettingData	= 43,
};


/* 
** Flow data types
**
** Used for val.flow.xlflow field of XLOPER and XLOPER12 structures
** when constructing flow-control XLOPERs and XLOPER12s
**/
enum XLOPER_FLOWTYPE
{
	xlflowHalt		= 1,
	xlflowGoto		= 2,
	xlflowRestart	= 8,
	xlflowPause		= 16,
	xlflowResume	= 64,
};


/*
** Return codes
**
** These values can be returned from Excel4(), Excel4v(), Excel12() or Excel12v().
*/
enum XLCALL_RESULT
{
	xlretSuccess		= 0,    /* success */
	xlretAbort			= 1,    /* macro halted */
	xlretInvXlfn		= 2,    /* invalid function number */
	xlretInvCount		= 4,    /* invalid number of arguments */
	xlretInvXloper		= 8,    /* invalid OPER structure */
	xlretStackOvfl		= 16,   /* stack overflow */
	xlretFailed			= 32,   /* command failed */
	xlretUncalced		= 64,   /* uncalced cell */
	xlretNotThreadSafe	= 128,  /* not allowed during multi-threaded calc */
	xlretInvAsynchronousContext	= 256,  /* invalid asynchronous function handle */
	xlretNotClusterSafe	= 512,  /* not supported on cluster */
};


/*
** XLL events
**
** Passed in to an xlEventRegister call to register a corresponding event.
*/

#define xleventCalculationEnded      1    /* Fires at the end of calculation */ 
#define xleventCalculationCanceled   2    /* Fires when calculation is interrupted */


/*
** Function prototypes
*/

#ifdef __cplusplus
extern "C" {
#endif

int _cdecl Excel4(int xlfn, LPXLOPER operRes, int count,... ); 
/* followed by count LPXLOPERs */

int pascal Excel4v(int xlfn, LPXLOPER operRes, int count, LPXLOPER opers[]);

int pascal XLCallVer(void);

long pascal LPenHelper(int wCode, VOID *lpv);

int _cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count,... );
/* followed by count LPXLOPER12s */

int pascal Excel12v(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[]);

#ifdef __cplusplus
}
#endif


/*
** Cluster Connector Async Callback
*/

typedef int (CALLBACK *PXL_HPC_ASYNC_CALLBACK)(LPXLOPER12 operAsyncHandle, LPXLOPER12 operReturn);


/*
** Cluster connector entry point return codes
*/

#define xlHpcRetSuccess            0
#define xlHpcRetSessionIdInvalid  -1
#define xlHpcRetCallFailed        -2


/* edit modes */
#define xlModeReady	0	// not in edit mode
#define xlModeEnter	1	// enter mode
#define xlModeEdit	2	// edit mode
#define xlModePoint	4	// point mode

/* document(page) types */
#define dtNil 0x7f	// window is not a sheet, macro, chart or basic
// OR window is not the selected window at idle state
#define dtSheet 0	// sheet
#define dtProc  1	// XLM macro
#define dtChart 2	// Chart
#define dtBasic 6	// VBA 

/* hit test codes */
#define htNone		0x00	// none of below
#define htClient	0x01	// internal for "in the client are", should never see
#define htVSplit	0x02	// vertical split area with split panes
#define htHSplit	0x03	// horizontal split area
#define htColWidth	0x04	// column width adjuster area
#define htRwHeight	0x05	// row height adjuster area
#define htRwColHdr	0x06	// the intersection of row and column headers
#define htObject	0x07	// the body of an object
// the following are for size handles of draw objects
#define htTopLeft	0x08
#define htBotLeft	0x09
#define htLeft		0x0A
#define htTopRight	0x0B
#define htBotRight	0x0C
#define htRight		0x0D
#define htTop		0x0E
#define htBot		0x0F
// end size handles
#define htRwGut		0x10	// row area of outline gutter
#define htColGut	0x11	// column area of outline gutter
#define htTextBox	0x12	// body of a text box (where we shouw I-Beam cursor)
#define htRwLevels	0x13	// row levels buttons of outline gutter
#define htColLevels	0x14	// column levels buttons of outline gutter
#define htDman		0x15	// the drag/drop handle of the selection
#define htDmanFill	0x16	// the auto-fill handle of the selection
#define htXSplit	0x17	// the intersection of the horz & vert pane splits
#define htVertex	0x18	// a vertex of a polygon draw object
#define htAddVtx	0x19	// htVertex in add a vertex mode
#define htDelVtx	0x1A	// htVertex in delete a vertex mode
#define htRwHdr		0x1B	// row header
#define htColHdr	0x1C	// column header
#define htRwShow	0x1D	// Like htRowHeight except means grow a hidden column
#define htColShow	0x1E	// column version of htRwShow
#define htSizing	0x1F	// Internal use only
#define htSxpivot	0x20	// a drag/drop tile in a pivot table
#define htTabs		0x21	// the sheet paging tabs
#define htEdit		0x22	// Internal use only

typedef struct _fmlainfo
{
	int wPointMode;	// current edit mode.  0 => rest of struct undefined
	int cch;	// count of characters in formula
	char *lpch;	// pointer to formula characters.  READ ONLY!!!
	int ichFirst;	// char offset to start of selection
	int ichLast;	// char offset to end of selection (may be > cch)
	int ichCaret;	// char offset to blinking caret
} FMLAINFO;

typedef struct _mouseinfo
{
	/* input section */
	HWND hwnd;		// window to get info on
	POINT pt;		// mouse position to get info on

	/* output section */
	int dt;			// document(page) type
	int ht;			// hit test code
	int rw;			// row @ mouse (-1 if #n/a)
	int col;		// col @ mouse (-1 if #n/a)
} MOUSEINFO;


/*
** Function number bits
*/
enum XLCALL_FUNCTYPE
{
	xlCommand = 0x8000,
	xlSpecial = 0x4000,
	xlIntl = 0x2000,
	xlPrompt = 0x1000,
};

//////////////////////////////////////////////////////////////////////////
// xlcall functions
//////////////////////////////////////////////////////////////////////////

#ifdef xlfAnd
#	undef xlfAnd
#endif // xlfAnd
#ifdef xlfOr
#	undef xlfOr
#endif // xlfOr
#ifdef xlfNot
#	undef xlfNot
#endif // xlfNot
#ifdef xlfMod
#	undef xlfMod
#endif // xlfMod
#ifdef xlfOffset
#	undef xlfOffset
#endif // xlfMod

namespace xlcall
{
	/*
	** Auxiliary function numbers
	**
	** These functions are available only from the C API,
	** not from the Excel macro language.
	*/

	enum XLCALL_SpecialFunc
	{
		xlFree = (0 | xlSpecial),
		xlStack = (1 | xlSpecial),
		xlCoerce = (2 | xlSpecial),
		xlSet = (3 | xlSpecial),
		xlSheetId = (4 | xlSpecial),
		xlSheetNm = (5 | xlSpecial),
		xlAbort = (6 | xlSpecial),
		xlGetInst = (7 | xlSpecial), /* Returns application's hinstance as an integer value, supported on 32-bit platform only */
		xlGetHwnd = (8 | xlSpecial),
		xlGetName = (9 | xlSpecial),
		xlEnableXLMsgs = (10 | xlSpecial),
		xlDisableXLMsgs = (11 | xlSpecial),
		xlDefineBinaryName = (12 | xlSpecial),
		xlGetBinaryName = (13 | xlSpecial),
		/* GetFooInfo are valid only for calls to LPenHelper */
		xlGetFmlaInfo = (14 | xlSpecial),
		xlGetMouseInfo = (15 | xlSpecial),
		xlAsyncReturn = (16 | xlSpecial),	/*Set return value from an asynchronous function call*/
		xlEventRegister = (17 | xlSpecial),	/*Register an XLL event*/
		xlRunningOnCluster = (18 | xlSpecial),	/*Returns true if running on Compute Cluster*/
		xlGetInstPtr = (19 | xlSpecial),	/* Returns application's hinstance as a handle, supported on both 32-bit and 64-bit platforms */
	};


	/*
	** User defined function
	**
	** First argument should be a function reference.
	*/
// #define xlUDF      255


	/*
	** Built-in Excel functions and command equivalents
	*/

	// Excel function numbers
	enum XLCALL_NormalFunc
	{
		xlfCount = 0,
		xlfIsna = 2,
		xlfIserror = 3,
		xlfSum = 4,
		xlfAverage = 5,
		xlfMin = 6,
		xlfMax = 7,
		xlfRow = 8,
		xlfColumn = 9,
		xlfNa = 10,
		xlfNpv = 11,
		xlfStdev = 12,
		xlfDollar = 13,
		xlfFixed = 14,
		xlfSin = 15,
		xlfCos = 16,
		xlfTan = 17,
		xlfAtan = 18,
		xlfPi = 19,
		xlfSqrt = 20,
		xlfExp = 21,
		xlfLn = 22,
		xlfLog10 = 23,
		xlfAbs = 24,
		xlfInt = 25,
		xlfSign = 26,
		xlfRound = 27,
		xlfLookup = 28,
		xlfIndex = 29,
		xlfRept = 30,
		xlfMid = 31,
		xlfLen = 32,
		xlfValue = 33,
		xlfTrue = 34,
		xlfFalse = 35,
		xlfAnd = 36,
		xlfOr = 37,
		xlfNot = 38,
		xlfMod = 39,
		xlfDcount = 40,
		xlfDsum = 41,
		xlfDaverage = 42,
		xlfDmin = 43,
		xlfDmax = 44,
		xlfDstdev = 45,
		xlfVar = 46,
		xlfDvar = 47,
		xlfText = 48,
		xlfLinest = 49,
		xlfTrend = 50,
		xlfLogest = 51,
		xlfGrowth = 52,
		xlfGoto = 53,
		xlfHalt = 54,
		xlfPv = 56,
		xlfFv = 57,
		xlfNper = 58,
		xlfPmt = 59,
		xlfRate = 60,
		xlfMirr = 61,
		xlfIrr = 62,
		xlfRand = 63,
		xlfMatch = 64,
		xlfDate = 65,
		xlfTime = 66,
		xlfDay = 67,
		xlfMonth = 68,
		xlfYear = 69,
		xlfWeekday = 70,
		xlfHour = 71,
		xlfMinute = 72,
		xlfSecond = 73,
		xlfNow = 74,
		xlfAreas = 75,
		xlfRows = 76,
		xlfColumns = 77,
		xlfOffset = 78,
		xlfAbsref = 79,
		xlfRelref = 80,
		xlfArgument = 81,
		xlfSearch = 82,
		xlfTranspose = 83,
		xlfError = 84,
		xlfStep = 85,
		xlfType = 86,
		xlfEcho = 87,
		xlfSetName = 88,
		xlfCaller = 89,
		xlfDeref = 90,
		xlfWindows = 91,
		xlfSeries = 92,
		xlfDocuments = 93,
		xlfActiveCell = 94,
		xlfSelection = 95,
		xlfResult = 96,
		xlfAtan2 = 97,
		xlfAsin = 98,
		xlfAcos = 99,
		xlfChoose = 100,
		xlfHlookup = 101,
		xlfVlookup = 102,
		xlfLinks = 103,
		xlfInput = 104,
		xlfIsref = 105,
		xlfGetFormula = 106,
		xlfGetName = 107,
		xlfSetValue = 108,
		xlfLog = 109,
		xlfExec = 110,
		xlfChar = 111,
		xlfLower = 112,
		xlfUpper = 113,
		xlfProper = 114,
		xlfLeft = 115,
		xlfRight = 116,
		xlfExact = 117,
		xlfTrim = 118,
		xlfReplace = 119,
		xlfSubstitute = 120,
		xlfCode = 121,
		xlfNames = 122,
		xlfDirectory = 123,
		xlfFind = 124,
		xlfCell = 125,
		xlfIserr = 126,
		xlfIstext = 127,
		xlfIsnumber = 128,
		xlfIsblank = 129,
		xlfT = 130,
		xlfN = 131,
		xlfFopen = 132,
		xlfFclose = 133,
		xlfFsize = 134,
		xlfFreadln = 135,
		xlfFread = 136,
		xlfFwriteln = 137,
		xlfFwrite = 138,
		xlfFpos = 139,
		xlfDatevalue = 140,
		xlfTimevalue = 141,
		xlfSln = 142,
		xlfSyd = 143,
		xlfDdb = 144,
		xlfGetDef = 145,
		xlfReftext = 146,
		xlfTextref = 147,
		xlfIndirect = 148,
		xlfRegister = 149,
		xlfCall = 150,
		xlfAddBar = 151,
		xlfAddMenu = 152,
		xlfAddCommand = 153,
		xlfEnableCommand = 154,
		xlfCheckCommand = 155,
		xlfRenameCommand = 156,
		xlfShowBar = 157,
		xlfDeleteMenu = 158,
		xlfDeleteCommand = 159,
		xlfGetChartItem = 160,
		xlfDialogBox = 161,
		xlfClean = 162,
		xlfMdeterm = 163,
		xlfMinverse = 164,
		xlfMmult = 165,
		xlfFiles = 166,
		xlfIpmt = 167,
		xlfPpmt = 168,
		xlfCounta = 169,
		xlfCancelKey = 170,
		xlfInitiate = 175,
		xlfRequest = 176,
		xlfPoke = 177,
		xlfExecute = 178,
		xlfTerminate = 179,
		xlfRestart = 180,
		xlfHelp = 181,
		xlfGetBar = 182,
		xlfProduct = 183,
		xlfFact = 184,
		xlfGetCell = 185,
		xlfGetWorkspace = 186,
		xlfGetWindow = 187,
		xlfGetDocument = 188,
		xlfDproduct = 189,
		xlfIsnontext = 190,
		xlfGetNote = 191,
		xlfNote = 192,
		xlfStdevp = 193,
		xlfVarp = 194,
		xlfDstdevp = 195,
		xlfDvarp = 196,
		xlfTrunc = 197,
		xlfIslogical = 198,
		xlfDcounta = 199,
		xlfDeleteBar = 200,
		xlfUnregister = 201,
		xlfUsdollar = 204,
		xlfFindb = 205,
		xlfSearchb = 206,
		xlfReplaceb = 207,
		xlfLeftb = 208,
		xlfRightb = 209,
		xlfMidb = 210,
		xlfLenb = 211,
		xlfRoundup = 212,
		xlfRounddown = 213,
		xlfAsc = 214,
		xlfDbcs = 215,
		xlfRank = 216,
		xlfAddress = 219,
		xlfDays360 = 220,
		xlfToday = 221,
		xlfVdb = 222,
		xlfMedian = 227,
		xlfSumproduct = 228,
		xlfSinh = 229,
		xlfCosh = 230,
		xlfTanh = 231,
		xlfAsinh = 232,
		xlfAcosh = 233,
		xlfAtanh = 234,
		xlfDget = 235,
		xlfCreateObject = 236,
		xlfVolatile = 237,
		xlfLastError = 238,
		xlfCustomUndo = 239,
		xlfCustomRepeat = 240,
		xlfFormulaConvert = 241,
		xlfGetLinkInfo = 242,
		xlfTextBox = 243,
		xlfInfo = 244,
		xlfGroup = 245,
		xlfGetObject = 246,
		xlfDb = 247,
		xlfPause = 248,
		xlfResume = 251,
		xlfFrequency = 252,
		xlfAddToolbar = 253,
		xlfDeleteToolbar = 254,
		xlUDF = 255,
		xlfResetToolbar = 256,
		xlfEvaluate = 257,
		xlfGetToolbar = 258,
		xlfGetTool = 259,
		xlfSpellingCheck = 260,
		xlfErrorType = 261,
		xlfAppTitle = 262,
		xlfWindowTitle = 263,
		xlfSaveToolbar = 264,
		xlfEnableTool = 265,
		xlfPressTool = 266,
		xlfRegisterId = 267,
		xlfGetWorkbook = 268,
		xlfAvedev = 269,
		xlfBetadist = 270,
		xlfGammaln = 271,
		xlfBetainv = 272,
		xlfBinomdist = 273,
		xlfChidist = 274,
		xlfChiinv = 275,
		xlfCombin = 276,
		xlfConfidence = 277,
		xlfCritbinom = 278,
		xlfEven = 279,
		xlfExpondist = 280,
		xlfFdist = 281,
		xlfFinv = 282,
		xlfFisher = 283,
		xlfFisherinv = 284,
		xlfFloor = 285,
		xlfGammadist = 286,
		xlfGammainv = 287,
		xlfCeiling = 288,
		xlfHypgeomdist = 289,
		xlfLognormdist = 290,
		xlfLoginv = 291,
		xlfNegbinomdist = 292,
		xlfNormdist = 293,
		xlfNormsdist = 294,
		xlfNorminv = 295,
		xlfNormsinv = 296,
		xlfStandardize = 297,
		xlfOdd = 298,
		xlfPermut = 299,
		xlfPoisson = 300,
		xlfTdist = 301,
		xlfWeibull = 302,
		xlfSumxmy2 = 303,
		xlfSumx2my2 = 304,
		xlfSumx2py2 = 305,
		xlfChitest = 306,
		xlfCorrel = 307,
		xlfCovar = 308,
		xlfForecast = 309,
		xlfFtest = 310,
		xlfIntercept = 311,
		xlfPearson = 312,
		xlfRsq = 313,
		xlfSteyx = 314,
		xlfSlope = 315,
		xlfTtest = 316,
		xlfProb = 317,
		xlfDevsq = 318,
		xlfGeomean = 319,
		xlfHarmean = 320,
		xlfSumsq = 321,
		xlfKurt = 322,
		xlfSkew = 323,
		xlfZtest = 324,
		xlfLarge = 325,
		xlfSmall = 326,
		xlfQuartile = 327,
		xlfPercentile = 328,
		xlfPercentrank = 329,
		xlfMode = 330,
		xlfTrimmean = 331,
		xlfTinv = 332,
		xlfMovieCommand = 334,
		xlfGetMovie = 335,
		xlfConcatenate = 336,
		xlfPower = 337,
		xlfPivotAddData = 338,
		xlfGetPivotTable = 339,
		xlfGetPivotField = 340,
		xlfGetPivotItem = 341,
		xlfRadians = 342,
		xlfDegrees = 343,
		xlfSubtotal = 344,
		xlfSumif = 345,
		xlfCountif = 346,
		xlfCountblank = 347,
		xlfScenarioGet = 348,
		xlfOptionsListsGet = 349,
		xlfIspmt = 350,
		xlfDatedif = 351,
		xlfDatestring = 352,
		xlfNumberstring = 353,
		xlfRoman = 354,
		xlfOpenDialog = 355,
		xlfSaveDialog = 356,
		xlfViewGet = 357,
		xlfGetpivotdata = 358,
		xlfHyperlink = 359,
		xlfPhonetic = 360,
		xlfAveragea = 361,
		xlfMaxa = 362,
		xlfMina = 363,
		xlfStdevpa = 364,
		xlfVarpa = 365,
		xlfStdeva = 366,
		xlfVara = 367,
		xlfBahttext = 368,
		xlfThaidayofweek = 369,
		xlfThaidigit = 370,
		xlfThaimonthofyear = 371,
		xlfThainumsound = 372,
		xlfThainumstring = 373,
		xlfThaistringlength = 374,
		xlfIsthaidigit = 375,
		xlfRoundbahtdown = 376,
		xlfRoundbahtup = 377,
		xlfThaiyear = 378,
		xlfRtd = 379,
		xlfCubevalue = 380,
		xlfCubemember = 381,
		xlfCubememberproperty = 382,
		xlfCuberankedmember = 383,
		xlfHex2bin = 384,
		xlfHex2dec = 385,
		xlfHex2oct = 386,
		xlfDec2bin = 387,
		xlfDec2hex = 388,
		xlfDec2oct = 389,
		xlfOct2bin = 390,
		xlfOct2hex = 391,
		xlfOct2dec = 392,
		xlfBin2dec = 393,
		xlfBin2oct = 394,
		xlfBin2hex = 395,
		xlfImsub = 396,
		xlfImdiv = 397,
		xlfImpower = 398,
		xlfImabs = 399,
		xlfImsqrt = 400,
		xlfImln = 401,
		xlfImlog2 = 402,
		xlfImlog10 = 403,
		xlfImsin = 404,
		xlfImcos = 405,
		xlfImexp = 406,
		xlfImargument = 407,
		xlfImconjugate = 408,
		xlfImaginary = 409,
		xlfImreal = 410,
		xlfComplex = 411,
		xlfImsum = 412,
		xlfImproduct = 413,
		xlfSeriessum = 414,
		xlfFactdouble = 415,
		xlfSqrtpi = 416,
		xlfQuotient = 417,
		xlfDelta = 418,
		xlfGestep = 419,
		xlfIseven = 420,
		xlfIsodd = 421,
		xlfMround = 422,
		xlfErf = 423,
		xlfErfc = 424,
		xlfBesselj = 425,
		xlfBesselk = 426,
		xlfBessely = 427,
		xlfBesseli = 428,
		xlfXirr = 429,
		xlfXnpv = 430,
		xlfPricemat = 431,
		xlfYieldmat = 432,
		xlfIntrate = 433,
		xlfReceived = 434,
		xlfDisc = 435,
		xlfPricedisc = 436,
		xlfYielddisc = 437,
		xlfTbilleq = 438,
		xlfTbillprice = 439,
		xlfTbillyield = 440,
		xlfPrice = 441,
		xlfYield = 442,
		xlfDollarde = 443,
		xlfDollarfr = 444,
		xlfNominal = 445,
		xlfEffect = 446,
		xlfCumprinc = 447,
		xlfCumipmt = 448,
		xlfEdate = 449,
		xlfEomonth = 450,
		xlfYearfrac = 451,
		xlfCoupdaybs = 452,
		xlfCoupdays = 453,
		xlfCoupdaysnc = 454,
		xlfCoupncd = 455,
		xlfCoupnum = 456,
		xlfCouppcd = 457,
		xlfDuration = 458,
		xlfMduration = 459,
		xlfOddlprice = 460,
		xlfOddlyield = 461,
		xlfOddfprice = 462,
		xlfOddfyield = 463,
		xlfRandbetween = 464,
		xlfWeeknum = 465,
		xlfAmordegrc = 466,
		xlfAmorlinc = 467,
		xlfConvert = 468,
		xlfAccrint = 469,
		xlfAccrintm = 470,
		xlfWorkday = 471,
		xlfNetworkdays = 472,
		xlfGcd = 473,
		xlfMultinomial = 474,
		xlfLcm = 475,
		xlfFvschedule = 476,
		xlfCubekpimember = 477,
		xlfCubeset = 478,
		xlfCubesetcount = 479,
		xlfIferror = 480,
		xlfCountifs = 481,
		xlfSumifs = 482,
		xlfAverageif = 483,
		xlfAverageifs = 484,
		xlfAggregate = 485,
		xlfBinom_dist = 486,
		xlfBinom_inv = 487,
		xlfConfidence_norm = 488,
		xlfConfidence_t = 489,
		xlfChisq_test = 490,
		xlfF_test = 491,
		xlfCovariance_p = 492,
		xlfCovariance_s = 493,
		xlfExpon_dist = 494,
		xlfGamma_dist = 495,
		xlfGamma_inv = 496,
		xlfMode_mult = 497,
		xlfMode_sngl = 498,
		xlfNorm_dist = 499,
		xlfNorm_inv = 500,
		xlfPercentile_exc = 501,
		xlfPercentile_inc = 502,
		xlfPercentrank_exc = 503,
		xlfPercentrank_inc = 504,
		xlfPoisson_dist = 505,
		xlfQuartile_exc = 506,
		xlfQuartile_inc = 507,
		xlfRank_avg = 508,
		xlfRank_eq = 509,
		xlfStdev_s = 510,
		xlfStdev_p = 511,
		xlfT_dist = 512,
		xlfT_dist_2t = 513,
		xlfT_dist_rt = 514,
		xlfT_inv = 515,
		xlfT_inv_2t = 516,
		xlfVar_s = 517,
		xlfVar_p = 518,
		xlfWeibull_dist = 519,
		xlfNetworkdays_intl = 520,
		xlfWorkday_intl = 521,
		xlfEcma_ceiling = 522,
		xlfIso_ceiling = 523,
		xlfBeta_dist = 525,
		xlfBeta_inv = 526,
		xlfChisq_dist = 527,
		xlfChisq_dist_rt = 528,
		xlfChisq_inv = 529,
		xlfChisq_inv_rt = 530,
		xlfF_dist = 531,
		xlfF_dist_rt = 532,
		xlfF_inv = 533,
		xlfF_inv_rt = 534,
		xlfHypgeom_dist = 535,
		xlfLognorm_dist = 536,
		xlfLognorm_inv = 537,
		xlfNegbinom_dist = 538,
		xlfNorm_s_dist = 539,
		xlfNorm_s_inv = 540,
		xlfT_test = 541,
		xlfZ_test = 542,
		xlfErf_precise = 543,
		xlfErfc_precise = 544,
		xlfGammaln_precise = 545,
		xlfCeiling_precise = 546,
		xlfFloor_precise = 547,
		xlfAcot = 548,
		xlfAcoth = 549,
		xlfCot = 550,
		xlfCoth = 551,
		xlfCsc = 552,
		xlfCsch = 553,
		xlfSec = 554,
		xlfSech = 555,
		xlfImtan = 556,
		xlfImcot = 557,
		xlfImcsc = 558,
		xlfImcsch = 559,
		xlfImsec = 560,
		xlfImsech = 561,
		xlfBitand = 562,
		xlfBitor = 563,
		xlfBitxor = 564,
		xlfBitlshift = 565,
		xlfBitrshift = 566,
		xlfPermutationa = 567,
		xlfCombina = 568,
		xlfXor = 569,
		xlfPduration = 570,
		xlfBase = 571,
		xlfDecimal = 572,
		xlfDays = 573,
		xlfBinom_dist_range = 574,
		xlfGamma = 575,
		xlfSkew_p = 576,
		xlfGauss = 577,
		xlfPhi = 578,
		xlfRri = 579,
		xlfUnichar = 580,
		xlfUnicode = 581,
		xlfMunit = 582,
		xlfArabic = 583,
		xlfIsoweeknum = 584,
		xlfNumbervalue = 585,
		xlfSheet = 586,
		xlfSheets = 587,
		xlfFormulatext = 588,
		xlfIsformula = 589,
		xlfIfna = 590,
		xlfCeiling_math = 591,
		xlfFloor_math = 592,
		xlfImsinh = 593,
		xlfImcosh = 594,
		xlfFilterxml = 595,
		xlfWebservice = 596,
		xlfEncodeurl = 597,
	};


	/* Excel command numbers */
	enum XLCALL_COMMAND
	{
		xlcBeep = (0 | xlCommand),
		xlcOpen = (1 | xlCommand),
		xlcOpenLinks = (2 | xlCommand),
		xlcCloseAll = (3 | xlCommand),
		xlcSave = (4 | xlCommand),
		xlcSaveAs = (5 | xlCommand),
		xlcFileDelete = (6 | xlCommand),
		xlcPageSetup = (7 | xlCommand),
		xlcPrint = (8 | xlCommand),
		xlcPrinterSetup = (9 | xlCommand),
		xlcQuit = (10 | xlCommand),
		xlcNewWindow = (11 | xlCommand),
		xlcArrangeAll = (12 | xlCommand),
		xlcWindowSize = (13 | xlCommand),
		xlcWindowMove = (14 | xlCommand),
		xlcFull = (15 | xlCommand),
		xlcClose = (16 | xlCommand),
		xlcRun = (17 | xlCommand),
		xlcSetPrintArea = (22 | xlCommand),
		xlcSetPrintTitles = (23 | xlCommand),
		xlcSetPageBreak = (24 | xlCommand),
		xlcRemovePageBreak = (25 | xlCommand),
		xlcFont = (26 | xlCommand),
		xlcDisplay = (27 | xlCommand),
		xlcProtectDocument = (28 | xlCommand),
		xlcPrecision = (29 | xlCommand),
		xlcA1R1c1 = (30 | xlCommand),
		xlcCalculateNow = (31 | xlCommand),
		xlcCalculation = (32 | xlCommand),
		xlcDataFind = (34 | xlCommand),
		xlcExtract = (35 | xlCommand),
		xlcDataDelete = (36 | xlCommand),
		xlcSetDatabase = (37 | xlCommand),
		xlcSetCriteria = (38 | xlCommand),
		xlcSort = (39 | xlCommand),
		xlcDataSeries = (40 | xlCommand),
		xlcTable = (41 | xlCommand),
		xlcFormatNumber = (42 | xlCommand),
		xlcAlignment = (43 | xlCommand),
		xlcStyle = (44 | xlCommand),
		xlcBorder = (45 | xlCommand),
		xlcCellProtection = (46 | xlCommand),
		xlcColumnWidth = (47 | xlCommand),
		xlcUndo = (48 | xlCommand),
		xlcCut = (49 | xlCommand),
		xlcCopy = (50 | xlCommand),
		xlcPaste = (51 | xlCommand),
		xlcClear = (52 | xlCommand),
		xlcPasteSpecial = (53 | xlCommand),
		xlcEditDelete = (54 | xlCommand),
		xlcInsert = (55 | xlCommand),
		xlcFillRight = (56 | xlCommand),
		xlcFillDown = (57 | xlCommand),
		xlcDefineName = (61 | xlCommand),
		xlcCreateNames = (62 | xlCommand),
		xlcFormulaGoto = (63 | xlCommand),
		xlcFormulaFind = (64 | xlCommand),
		xlcSelectLastCell = (65 | xlCommand),
		xlcShowActiveCell = (66 | xlCommand),
		xlcGalleryArea = (67 | xlCommand),
		xlcGalleryBar = (68 | xlCommand),
		xlcGalleryColumn = (69 | xlCommand),
		xlcGalleryLine = (70 | xlCommand),
		xlcGalleryPie = (71 | xlCommand),
		xlcGalleryScatter = (72 | xlCommand),
		xlcCombination = (73 | xlCommand),
		xlcPreferred = (74 | xlCommand),
		xlcAddOverlay = (75 | xlCommand),
		xlcGridlines = (76 | xlCommand),
		xlcSetPreferred = (77 | xlCommand),
		xlcAxes = (78 | xlCommand),
		xlcLegend = (79 | xlCommand),
		xlcAttachText = (80 | xlCommand),
		xlcAddArrow = (81 | xlCommand),
		xlcSelectChart = (82 | xlCommand),
		xlcSelectPlotArea = (83 | xlCommand),
		xlcPatterns = (84 | xlCommand),
		xlcMainChart = (85 | xlCommand),
		xlcOverlay = (86 | xlCommand),
		xlcScale = (87 | xlCommand),
		xlcFormatLegend = (88 | xlCommand),
		xlcFormatText = (89 | xlCommand),
		xlcEditRepeat = (90 | xlCommand),
		xlcParse = (91 | xlCommand),
		xlcJustify = (92 | xlCommand),
		xlcHide = (93 | xlCommand),
		xlcUnhide = (94 | xlCommand),
		xlcWorkspace = (95 | xlCommand),
		xlcFormula = (96 | xlCommand),
		xlcFormulaFill = (97 | xlCommand),
		xlcFormulaArray = (98 | xlCommand),
		xlcDataFindNext = (99 | xlCommand),
		xlcDataFindPrev = (100 | xlCommand),
		xlcFormulaFindNext = (101 | xlCommand),
		xlcFormulaFindPrev = (102 | xlCommand),
		xlcActivate = (103 | xlCommand),
		xlcActivateNext = (104 | xlCommand),
		xlcActivatePrev = (105 | xlCommand),
		xlcUnlockedNext = (106 | xlCommand),
		xlcUnlockedPrev = (107 | xlCommand),
		xlcCopyPicture = (108 | xlCommand),
		xlcSelect = (109 | xlCommand),
		xlcDeleteName = (110 | xlCommand),
		xlcDeleteFormat = (111 | xlCommand),
		xlcVline = (112 | xlCommand),
		xlcHline = (113 | xlCommand),
		xlcVpage = (114 | xlCommand),
		xlcHpage = (115 | xlCommand),
		xlcVscroll = (116 | xlCommand),
		xlcHscroll = (117 | xlCommand),
		xlcAlert = (118 | xlCommand),
		xlcNew = (119 | xlCommand),
		xlcCancelCopy = (120 | xlCommand),
		xlcShowClipboard = (121 | xlCommand),
		xlcMessage = (122 | xlCommand),
		xlcPasteLink = (124 | xlCommand),
		xlcAppActivate = (125 | xlCommand),
		xlcDeleteArrow = (126 | xlCommand),
		xlcRowHeight = (127 | xlCommand),
		xlcFormatMove = (128 | xlCommand),
		xlcFormatSize = (129 | xlCommand),
		xlcFormulaReplace = (130 | xlCommand),
		xlcSendKeys = (131 | xlCommand),
		xlcSelectSpecial = (132 | xlCommand),
		xlcApplyNames = (133 | xlCommand),
		xlcReplaceFont = (134 | xlCommand),
		xlcFreezePanes = (135 | xlCommand),
		xlcShowInfo = (136 | xlCommand),
		xlcSplit = (137 | xlCommand),
		xlcOnWindow = (138 | xlCommand),
		xlcOnData = (139 | xlCommand),
		xlcDisableInput = (140 | xlCommand),
		xlcEcho = (141 | xlCommand),
		xlcOutline = (142 | xlCommand),
		xlcListNames = (143 | xlCommand),
		xlcFileClose = (144 | xlCommand),
		xlcSaveWorkbook = (145 | xlCommand),
		xlcDataForm = (146 | xlCommand),
		xlcCopyChart = (147 | xlCommand),
		xlcOnTime = (148 | xlCommand),
		xlcWait = (149 | xlCommand),
		xlcFormatFont = (150 | xlCommand),
		xlcFillUp = (151 | xlCommand),
		xlcFillLeft = (152 | xlCommand),
		xlcDeleteOverlay = (153 | xlCommand),
		xlcNote = (154 | xlCommand),
		xlcShortMenus = (155 | xlCommand),
		xlcSetUpdateStatus = (159 | xlCommand),
		xlcColorPalette = (161 | xlCommand),
		xlcDeleteStyle = (162 | xlCommand),
		xlcWindowRestore = (163 | xlCommand),
		xlcWindowMaximize = (164 | xlCommand),
		xlcError = (165 | xlCommand),
		xlcChangeLink = (166 | xlCommand),
		xlcCalculateDocument = (167 | xlCommand),
		xlcOnKey = (168 | xlCommand),
		xlcAppRestore = (169 | xlCommand),
		xlcAppMove = (170 | xlCommand),
		xlcAppSize = (171 | xlCommand),
		xlcAppMinimize = (172 | xlCommand),
		xlcAppMaximize = (173 | xlCommand),
		xlcBringToFront = (174 | xlCommand),
		xlcSendToBack = (175 | xlCommand),
		xlcMainChartType = (185 | xlCommand),
		xlcOverlayChartType = (186 | xlCommand),
		xlcSelectEnd = (187 | xlCommand),
		xlcOpenMail = (188 | xlCommand),
		xlcSendMail = (189 | xlCommand),
		xlcStandardFont = (190 | xlCommand),
		xlcConsolidate = (191 | xlCommand),
		xlcSortSpecial = (192 | xlCommand),
		xlcGallery3dArea = (193 | xlCommand),
		xlcGallery3dColumn = (194 | xlCommand),
		xlcGallery3dLine = (195 | xlCommand),
		xlcGallery3dPie = (196 | xlCommand),
		xlcView3d = (197 | xlCommand),
		xlcGoalSeek = (198 | xlCommand),
		xlcWorkgroup = (199 | xlCommand),
		xlcFillGroup = (200 | xlCommand),
		xlcUpdateLink = (201 | xlCommand),
		xlcPromote = (202 | xlCommand),
		xlcDemote = (203 | xlCommand),
		xlcShowDetail = (204 | xlCommand),
		xlcUngroup = (206 | xlCommand),
		xlcObjectProperties = (207 | xlCommand),
		xlcSaveNewObject = (208 | xlCommand),
		xlcShare = (209 | xlCommand),
		xlcShareName = (210 | xlCommand),
		xlcDuplicate = (211 | xlCommand),
		xlcApplyStyle = (212 | xlCommand),
		xlcAssignToObject = (213 | xlCommand),
		xlcObjectProtection = (214 | xlCommand),
		xlcHideObject = (215 | xlCommand),
		xlcSetExtract = (216 | xlCommand),
		xlcCreatePublisher = (217 | xlCommand),
		xlcSubscribeTo = (218 | xlCommand),
		xlcAttributes = (219 | xlCommand),
		xlcShowToolbar = (220 | xlCommand),
		xlcPrintPreview = (222 | xlCommand),
		xlcEditColor = (223 | xlCommand),
		xlcShowLevels = (224 | xlCommand),
		xlcFormatMain = (225 | xlCommand),
		xlcFormatOverlay = (226 | xlCommand),
		xlcOnRecalc = (227 | xlCommand),
		xlcEditSeries = (228 | xlCommand),
		xlcDefineStyle = (229 | xlCommand),
		xlcLinePrint = (240 | xlCommand),
		xlcEnterData = (243 | xlCommand),
		xlcGalleryRadar = (249 | xlCommand),
		xlcMergeStyles = (250 | xlCommand),
		xlcEditionOptions = (251 | xlCommand),
		xlcPastePicture = (252 | xlCommand),
		xlcPastePictureLink = (253 | xlCommand),
		xlcSpelling = (254 | xlCommand),
		xlcZoom = (256 | xlCommand),
		xlcResume = (258 | xlCommand),
		xlcInsertObject = (259 | xlCommand),
		xlcWindowMinimize = (260 | xlCommand),
		xlcSize = (261 | xlCommand),
		xlcMove = (262 | xlCommand),
		xlcSoundNote = (265 | xlCommand),
		xlcSoundPlay = (266 | xlCommand),
		xlcFormatShape = (267 | xlCommand),
		xlcExtendPolygon = (268 | xlCommand),
		xlcFormatAuto = (269 | xlCommand),
		xlcGallery3dBar = (272 | xlCommand),
		xlcGallery3dSurface = (273 | xlCommand),
		xlcFillAuto = (274 | xlCommand),
		xlcCustomizeToolbar = (276 | xlCommand),
		xlcAddTool = (277 | xlCommand),
		xlcEditObject = (278 | xlCommand),
		xlcOnDoubleclick = (279 | xlCommand),
		xlcOnEntry = (280 | xlCommand),
		xlcWorkbookAdd = (281 | xlCommand),
		xlcWorkbookMove = (282 | xlCommand),
		xlcWorkbookCopy = (283 | xlCommand),
		xlcWorkbookOptions = (284 | xlCommand),
		xlcSaveWorkspace = (285 | xlCommand),
		xlcChartWizard = (288 | xlCommand),
		xlcDeleteTool = (289 | xlCommand),
		xlcMoveTool = (290 | xlCommand),
		xlcWorkbookSelect = (291 | xlCommand),
		xlcWorkbookActivate = (292 | xlCommand),
		xlcAssignToTool = (293 | xlCommand),
		xlcCopyTool = (295 | xlCommand),
		xlcResetTool = (296 | xlCommand),
		xlcConstrainNumeric = (297 | xlCommand),
		xlcPasteTool = (298 | xlCommand),
		xlcPlacement = (300 | xlCommand),
		xlcFillWorkgroup = (301 | xlCommand),
		xlcWorkbookNew = (302 | xlCommand),
		xlcScenarioCells = (305 | xlCommand),
		xlcScenarioDelete = (306 | xlCommand),
		xlcScenarioAdd = (307 | xlCommand),
		xlcScenarioEdit = (308 | xlCommand),
		xlcScenarioShow = (309 | xlCommand),
		xlcScenarioShowNext = (310 | xlCommand),
		xlcScenarioSummary = (311 | xlCommand),
		xlcPivotTableWizard = (312 | xlCommand),
		xlcPivotFieldProperties = (313 | xlCommand),
		xlcPivotField = (314 | xlCommand),
		xlcPivotItem = (315 | xlCommand),
		xlcPivotAddFields = (316 | xlCommand),
		xlcOptionsCalculation = (318 | xlCommand),
		xlcOptionsEdit = (319 | xlCommand),
		xlcOptionsView = (320 | xlCommand),
		xlcAddinManager = (321 | xlCommand),
		xlcMenuEditor = (322 | xlCommand),
		xlcAttachToolbars = (323 | xlCommand),
		xlcVbaactivate = (324 | xlCommand),
		xlcOptionsChart = (325 | xlCommand),
		xlcVbaInsertFile = (328 | xlCommand),
		xlcVbaProcedureDefinition = (330 | xlCommand),
		xlcRoutingSlip = (336 | xlCommand),
		xlcRouteDocument = (338 | xlCommand),
		xlcMailLogon = (339 | xlCommand),
		xlcInsertPicture = (342 | xlCommand),
		xlcEditTool = (343 | xlCommand),
		xlcGalleryDoughnut = (344 | xlCommand),
		xlcChartTrend = (350 | xlCommand),
		xlcPivotItemProperties = (352 | xlCommand),
		xlcWorkbookInsert = (354 | xlCommand),
		xlcOptionsTransition = (355 | xlCommand),
		xlcOptionsGeneral = (356 | xlCommand),
		xlcFilterAdvanced = (370 | xlCommand),
		xlcMailAddMailer = (373 | xlCommand),
		xlcMailDeleteMailer = (374 | xlCommand),
		xlcMailReply = (375 | xlCommand),
		xlcMailReplyAll = (376 | xlCommand),
		xlcMailForward = (377 | xlCommand),
		xlcMailNextLetter = (378 | xlCommand),
		xlcDataLabel = (379 | xlCommand),
		xlcInsertTitle = (380 | xlCommand),
		xlcFontProperties = (381 | xlCommand),
		xlcMacroOptions = (382 | xlCommand),
		xlcWorkbookHide = (383 | xlCommand),
		xlcWorkbookUnhide = (384 | xlCommand),
		xlcWorkbookDelete = (385 | xlCommand),
		xlcWorkbookName = (386 | xlCommand),
		xlcGalleryCustom = (388 | xlCommand),
		xlcAddChartAutoformat = (390 | xlCommand),
		xlcDeleteChartAutoformat = (391 | xlCommand),
		xlcChartAddData = (392 | xlCommand),
		xlcAutoOutline = (393 | xlCommand),
		xlcTabOrder = (394 | xlCommand),
		xlcShowDialog = (395 | xlCommand),
		xlcSelectAll = (396 | xlCommand),
		xlcUngroupSheets = (397 | xlCommand),
		xlcSubtotalCreate = (398 | xlCommand),
		xlcSubtotalRemove = (399 | xlCommand),
		xlcRenameObject = (400 | xlCommand),
		xlcWorkbookScroll = (412 | xlCommand),
		xlcWorkbookNext = (413 | xlCommand),
		xlcWorkbookPrev = (414 | xlCommand),
		xlcWorkbookTabSplit = (415 | xlCommand),
		xlcFullScreen = (416 | xlCommand),
		xlcWorkbookProtect = (417 | xlCommand),
		xlcScrollbarProperties = (420 | xlCommand),
		xlcPivotShowPages = (421 | xlCommand),
		xlcTextToColumns = (422 | xlCommand),
		xlcFormatCharttype = (423 | xlCommand),
		xlcLinkFormat = (424 | xlCommand),
		xlcTracerDisplay = (425 | xlCommand),
		xlcTracerNavigate = (430 | xlCommand),
		xlcTracerClear = (431 | xlCommand),
		xlcTracerError = (432 | xlCommand),
		xlcPivotFieldGroup = (433 | xlCommand),
		xlcPivotFieldUngroup = (434 | xlCommand),
		xlcCheckboxProperties = (435 | xlCommand),
		xlcLabelProperties = (436 | xlCommand),
		xlcListboxProperties = (437 | xlCommand),
		xlcEditboxProperties = (438 | xlCommand),
		xlcPivotRefresh = (439 | xlCommand),
		xlcLinkCombo = (440 | xlCommand),
		xlcOpenText = (441 | xlCommand),
		xlcHideDialog = (442 | xlCommand),
		xlcSetDialogFocus = (443 | xlCommand),
		xlcEnableObject = (444 | xlCommand),
		xlcPushbuttonProperties = (445 | xlCommand),
		xlcSetDialogDefault = (446 | xlCommand),
		xlcFilter = (447 | xlCommand),
		xlcFilterShowAll = (448 | xlCommand),
		xlcClearOutline = (449 | xlCommand),
		xlcFunctionWizard = (450 | xlCommand),
		xlcAddListItem = (451 | xlCommand),
		xlcSetListItem = (452 | xlCommand),
		xlcRemoveListItem = (453 | xlCommand),
		xlcSelectListItem = (454 | xlCommand),
		xlcSetControlValue = (455 | xlCommand),
		xlcSaveCopyAs = (456 | xlCommand),
		xlcOptionsListsAdd = (458 | xlCommand),
		xlcOptionsListsDelete = (459 | xlCommand),
		xlcSeriesAxes = (460 | xlCommand),
		xlcSeriesX = (461 | xlCommand),
		xlcSeriesY = (462 | xlCommand),
		xlcErrorbarX = (463 | xlCommand),
		xlcErrorbarY = (464 | xlCommand),
		xlcFormatChart = (465 | xlCommand),
		xlcSeriesOrder = (466 | xlCommand),
		xlcMailLogoff = (467 | xlCommand),
		xlcClearRoutingSlip = (468 | xlCommand),
		xlcAppActivateMicrosoft = (469 | xlCommand),
		xlcMailEditMailer = (470 | xlCommand),
		xlcOnSheet = (471 | xlCommand),
		xlcStandardWidth = (472 | xlCommand),
		xlcScenarioMerge = (473 | xlCommand),
		xlcSummaryInfo = (474 | xlCommand),
		xlcFindFile = (475 | xlCommand),
		xlcActiveCellFont = (476 | xlCommand),
		xlcEnableTipwizard = (477 | xlCommand),
		xlcVbaMakeAddin = (478 | xlCommand),
		xlcInsertdatatable = (480 | xlCommand),
		xlcWorkgroupOptions = (481 | xlCommand),
		xlcMailSendMailer = (482 | xlCommand),
		xlcAutocorrect = (485 | xlCommand),
		xlcPostDocument = (489 | xlCommand),
		xlcPicklist = (491 | xlCommand),
		xlcViewShow = (493 | xlCommand),
		xlcViewDefine = (494 | xlCommand),
		xlcViewDelete = (495 | xlCommand),
		xlcSheetBackground = (509 | xlCommand),
		xlcInsertMapObject = (510 | xlCommand),
		xlcOptionsMenono = (511 | xlCommand),
		xlcNormal = (518 | xlCommand),
		xlcLayout = (519 | xlCommand),
		xlcRmPrintArea = (520 | xlCommand),
		xlcClearPrintArea = (521 | xlCommand),
		xlcAddPrintArea = (522 | xlCommand),
		xlcMoveBrk = (523 | xlCommand),
		xlcHidecurrNote = (545 | xlCommand),
		xlcHideallNotes = (546 | xlCommand),
		xlcDeleteNote = (547 | xlCommand),
		xlcTraverseNotes = (548 | xlCommand),
		xlcActivateNotes = (549 | xlCommand),
		xlcProtectRevisions = (620 | xlCommand),
		xlcUnprotectRevisions = (621 | xlCommand),
		xlcOptionsMe = (647 | xlCommand),
		xlcWebPublish = (653 | xlCommand),
		xlcNewwebquery = (667 | xlCommand),
		xlcPivotTableChart = (673 | xlCommand),
		xlcOptionsSave = (753 | xlCommand),
		xlcOptionsSpell = (755 | xlCommand),
		xlcHideallInkannots = (808 | xlCommand),
	};
};
#endif // __XLCALL_H__
