/*
'    Doom Builder
'    Copyright (c) 2003 Pascal vd Heiden, www.codeimp.com
'    This program is released under GNU General Public License
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
*/


// Definitions
#define WIN32_LEAN_AND_MEAN
#define CLIPBUFFER_RESOLUTION		1280				//3200
#define CLIPBUFFER_HALFRESOLUTION	640				//1600
#define CLIPBUFFER_MULTIPLIER		203.71832716f		//509.2958179f

// Includes
#include <windows.h>
#include <objbase.h>
#include <math.h>
#include "clip.h"

// Variables
unsigned char clipbuffer[CLIPBUFFER_RESOLUTION];


// InitializeClipper: Initializes clipper
//-----------------------------------------------------------------------------
void __fastcall InitializeClipper()
{
	// Erase the clipbuffer
	memset(clipbuffer, 0, CLIPBUFFER_RESOLUTION);
}


// TerminateClipper: Terminates clipper
//-----------------------------------------------------------------------------
void __fastcall TerminateClipper()
{
}


// TestClipRange: Tests if anything between 2 given angels is not clipped.
//-----------------------------------------------------------------------------
int __fastcall TestClipRange(float a1, float a2)
{
	// Calculate the buffer coordinates
	int c1 = (int)(a1 * CLIPBUFFER_MULTIPLIER);
	int c2 = (int)(a2 * CLIPBUFFER_MULTIPLIER);
	while(c1 < 0) c1 += CLIPBUFFER_RESOLUTION;
	while(c2 < 0) c2 += CLIPBUFFER_RESOLUTION;
	while(c1 >= CLIPBUFFER_RESOLUTION) c1 -= CLIPBUFFER_RESOLUTION;
	while(c2 >= CLIPBUFFER_RESOLUTION) c2 -= CLIPBUFFER_RESOLUTION;
	
	// Determine rotation direction
	int m = c2 - c1;
	if(m < 0) m += CLIPBUFFER_RESOLUTION;
	if(m < CLIPBUFFER_HALFRESOLUTION)
	{
		// Check if the range goes through zero point
		if(c2 < c1)
		{
			// Test two partitions
			return (memchr(clipbuffer + c1, 0, CLIPBUFFER_RESOLUTION - c1) || memchr(clipbuffer, 0, c2 + 1)) != NULL;
		}
		else
		{
			// Test entire stroke
			return memchr(clipbuffer + c1, 0, c2 - c1 + 1) != NULL;
		}
	}
	else
	{
		// Check if the range goes through zero point
		if(c2 > c1)
		{
			// Test two partitions
			return (memchr(clipbuffer, 0, c1 + 1) || memchr(clipbuffer + c2, 0, CLIPBUFFER_RESOLUTION - c2)) != NULL;
		}
		else
		{
			// Test entire stroke
			return memchr(clipbuffer + c2, 0, c1 - c2 + 1) != NULL;
		}
	}
}


// WriteClipRange: Writes a range on the clipbuffer.
//-----------------------------------------------------------------------------
void __fastcall WriteClipRange(float a1, float a2)
{
	// Calculate the buffer coordinates
	int c1 = (int)(a1 * CLIPBUFFER_MULTIPLIER);
	int c2 = (int)(a2 * CLIPBUFFER_MULTIPLIER);
	while(c1 < 0) c1 += CLIPBUFFER_RESOLUTION;
	while(c2 < 0) c2 += CLIPBUFFER_RESOLUTION;
	while(c1 >= CLIPBUFFER_RESOLUTION) c1 -= CLIPBUFFER_RESOLUTION;
	while(c2 >= CLIPBUFFER_RESOLUTION) c2 -= CLIPBUFFER_RESOLUTION;
	
	// Start at the beginning angle
	int i = c1;
	
	// Determine rotation direction
	int m = c2 - c1;
	if(m < 0) m += CLIPBUFFER_RESOLUTION;
	if(m < CLIPBUFFER_HALFRESOLUTION)
	{
		// Check if the range goes through zero point
		if(c2 < c1)
		{
			// Write in two partitions
			memset(clipbuffer + c1, 1, CLIPBUFFER_RESOLUTION - c1);
			memset(clipbuffer, 1, c2 + 1);
		}
		else
		{
			// Write entire stroke
			memset(clipbuffer + c1, 1, c2 - c1 + 1);
		}
	}
	else
	{
		// Check if the range goes through zero point
		if(c2 > c1)
		{
			// Write in two partitions
			memset(clipbuffer, 1, c1 + 1);
			memset(clipbuffer + c2, 1, CLIPBUFFER_RESOLUTION - c2);
		}
		else
		{
			// Write entire stroke
			memset(clipbuffer + c2, 1, c1 - c2 + 1);
		}
	}
}


// TestFullClipbuffer: Tests if the entire clipbuffer has been set.
// returns true when entire buffer set, otherwise returns false.
//-----------------------------------------------------------------------------
int __fastcall TestFullClipbuffer()
{
	// Check if there is any zero byte and return the result
	return memchr(clipbuffer, 0, CLIPBUFFER_RESOLUTION) == NULL;
}

