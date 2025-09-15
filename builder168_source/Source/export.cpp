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

// Includes
#include <windows.h>
#include <objbase.h>
#include "data.h"
#include "math.h"


// ExportWavefrontObj: Exports the entire map as Wavefront Obj
//-----------------------------------------------------------------------------
void APIENTRY ExportWavefrontObj(char* filepathname,
								 MAPVERTEX* vertices, MAPLINEDEF* linedefs,
								 MAPSIDEDEF* sidedefs, MAPSECTOR* sectors,
								 MAPTHING* things,
								 int numvertices, int numlinedefs,
								 int numsidedefs, int numsectors,
								 int numthings)
{
	
	// TODO: Export the entire map as Wavefront Obj, write it to 'filepathname'
	
}

