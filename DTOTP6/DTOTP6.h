/* +++Date last modified: 05-Jul-1997 */
/*
$Header: c:/bnxl/rcs/dtotp6.h 1.1 1995/05/01 00:04:08 bnelson Exp $

$Log: dtotp6.h $
Revision 1.1  1995/05/01 00:04:08  bnelson

- Include file companion to dtotp6.c, see notes contained there.
*/

// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the DTOTP6_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// DTOTP6_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef DTOTP6_EXPORTS
#define DTOTP6_API __declspec(dllexport)
#else
#define DTOTP6_API __declspec(dllimport)
#endif

// This class is exported from the DTOTP6.dll
class DTOTP6_API CDTOTP6 {
public:
	CDTOTP6(void);
	// TODO: add your methods here.
};

extern DTOTP6_API int nDTOTP6;

//#include "dirport.h"

#ifndef D2TOTP6_H_
#define     D2TOTP6_H_

//#ifdef __TURBOC__
//#pragma     option -a-       /* Force byte alignment in struct */
//#endif

//#ifdef __GNUC__
//#define PAK        __attribute__((packed))
//#else
#define PAK
//#endif

#ifdef MONOSPACE_6           /* Just to be safe... */
#define     double_to_tp6    DBL2TP
#define     tp6_to_double    TP2DBL
#endif

typedef struct {
    unsigned char be   PAK;     /* biased exponent */
    unsigned int  v1   PAK;     /* lower 16 bits of mantissa */
    unsigned int  v2   PAK;     /* next  16 bits of mantissa */
    unsigned int  v3:7 PAK;     /* upper  7 bits of mantissa */
    unsigned int  s :1 PAK;     /* sign bit */
} tp_real_t;

extern DTOTP6_API void *DtoTP6(double x, void *TP6buffer);
extern DTOTP6_API double TP6toD(void *TP6buffer);

#endif    /* D2TOTP6_H_ */

