
#define SET_Hook_XLL 0
#define SET_Hook_XLLExport 0
#define SET_Hook_Other 1

extern void AttachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz);
extern void DetachFunction(PVOID *ppvReal, PVOID pvMine, PCHAR psz);
