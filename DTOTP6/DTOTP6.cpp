/* +++Date last modified: 05-Jul-1997 */
/*
$Header: c:/bnxl/rcs/dtotp6.c 1.2 1995/05/01 00:30:58 bnelson Exp $

$Log: dtotp6.c $
Revision 1.2  1995/05/01 00:30:58  bnelson

- Added test driver and Thad Smith's original function
  to convert a TP real to double

- Checks out OK with lint.

- Tested on GNU C 2.6.3 (little endian Intel) after adding PAK to
  real struct.

Revision 1.1  1995/04/30 23:54:56  bnelson

Written by Bob Nelson of Dallas, TX, USA (bnelson@netcom.com)

Original tp6_to_double() written by Thad Smith III of Boulder, CO, and
  released to the public domain in SNIPPETS

- Initial release -- converts C double value into the bit pattern used
  by a Turbo Pascal 6-byte real. Uses the "real" struct written by Thad
  Smith for ease of assignment to members.

- Tested on BC++ 3.1. 

- This source and associated include are contributed to the Public Domain.
*/

#include "stdafx.h"
#include <math.h>
#include "DTOTP6.h"

#define DBL_BIAS            0x3FE
#define REAL_BIAS           0x80
#define TP_REAL_BIAS        (DBL_BIAS - REAL_BIAS)    /* 0x37E */

BOOL APIENTRY DllMain( HANDLE hModule, 
					   DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}


// This is an example of an exported variable
DTOTP6_API int nDTOTP6=0;

// This is the constructor of a class that has been exported.
// see DTOTP6.h for the class definition
CDTOTP6::CDTOTP6()
{ 
	return; 
}
// This is an example of an exported function.
DTOTP6_API void *DtoTP6(double x, void *TP6buffer)
{
	unsigned short int *wp;
	tp_real_t r;

	if(x == 0.0)
	{
		r.v3 = r.v2 = r.v1 = r.be = r.s = 0;
		return memcpy(TP6buffer, &r, 6);
	}

	wp = (unsigned short int *)&x;      // Break down double into words
	r.s  = wp[3] >> 15;					// High bit set for sign

	// ------------------------------------------------------------------
	// Grab biased exponent -- exclude sign and shift out the MSB
	// mantissa bits.

	r.be = (unsigned char)(((wp[3] & 0x7FFF) >> 4) - TP_REAL_BIAS);

	// ------------------------------------------------------------------
	// Now...just assign the mantissa after shifting the bits to conform
	// with the layout for the TP 6-byte real.

	r.v3 = ((wp[3] & 0x0F) << 3) | (wp[2] >> 13);
	r.v2 = (wp[2] << 3) | (wp[1] >> 13);
	r.v1 = (wp[1] << 3) | (wp[0] >> 13);
	
	return memcpy(TP6buffer, &r, 6);
}

// -----------------------------------------------------------------
// Slightly adapted version of Thad Smith's function from TP6TOD.C
// from Snippets. (Uses TP real struct parameter and no memcpy).

DTOTP6_API double TP6toD(void *TP6buffer)
{
	tp_real_t r;
	memcpy(&r, TP6buffer, 6);
	if (r.be == 0)
        return 0.0;

	return ((((128 + r.v3) * 65536.0) + r.v2) * 65536.0 + r.v1) * 
		ldexp((r.s ? -1.0 : 1.0), r.be - (129 + 39));
}



