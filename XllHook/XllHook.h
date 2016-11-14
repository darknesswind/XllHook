
#define SET_Hook_XLL 1
#if SET_Hook_XLL
#	define SET_Hook_XLLExport 0
#else
#	define SET_Hook_XLLExport 0
#endif
#define SET_Hook_Other 0

extern void AttachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz);
extern void DetachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz);
