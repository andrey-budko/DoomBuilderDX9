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

// Includes
#include <windows.h>
#include <objbase.h>
#include <stdio.h>
#include <wchar.h>
#include "data.h"
#include "const.h"

// Limits
int maxbackups;
const int BACKUPLIMIT = 102;

// Structures backup
MAPTHING* backup_things[BACKUPLIMIT];
MAPLINEDEF* backup_linedefs[BACKUPLIMIT];
MAPSIDEDEF* backup_sidedefs[BACKUPLIMIT];
MAPVERTEX* backup_vertices[BACKUPLIMIT];
MAPSECTOR* backup_sectors[BACKUPLIMIT];
int backup_numthings[BACKUPLIMIT];
int backup_numlinedefs[BACKUPLIMIT];
int backup_numsidedefs[BACKUPLIMIT];
int backup_numvertices[BACKUPLIMIT];
int backup_numsectors[BACKUPLIMIT];


// UndoRedo_Init: Initializes pointers
//----------------------------------------------------------------------------
void APIENTRY UndoRedo_Init()
{
	// Go for all backup levels to deallocate
	for(int i = 0; i < BACKUPLIMIT; i++)
	{
		// Zero pointers and numbers
		backup_things[i] = NULL;
		backup_linedefs[i] = NULL;
		backup_sidedefs[i] = NULL;
		backup_vertices[i] = NULL;
		backup_sectors[i] = NULL;
		backup_numthings[i] = 0;
		backup_numlinedefs[i] = 0;
		backup_numsidedefs[i] = 0;
		backup_numvertices[i] = 0;
		backup_numsectors[i] = 0;
	}
}


// UndoRedo_Term: Deallocates any allocated backup memory
//----------------------------------------------------------------------------
void APIENTRY UndoRedo_Term()
{
	// Go for all allocated backup levels to deallocate
	for(int i = 0; i < BACKUPLIMIT; i++)
	{
		// Deallocate if set
		if(backup_things[i])
		{
			delete[] backup_things[i];
			backup_things[i] = NULL;
			backup_numthings[i] = 0;
		}
		if(backup_linedefs[i])
		{
			delete[] backup_linedefs[i];
			backup_linedefs[i] = NULL;
			backup_numlinedefs[i] = 0;
		}
		if(backup_sidedefs[i]) 
		{
			delete[] backup_sidedefs[i];
			backup_sidedefs[i] = NULL;
			backup_numsidedefs[i] = 0;
		}
		if(backup_vertices[i])
		{
			delete[] backup_vertices[i];
			backup_vertices[i] = NULL;
			backup_numvertices[i] = 0;
		}
		if(backup_sectors[i])
		{
			delete[] backup_sectors[i];
			backup_sectors[i] = NULL;
			backup_numsectors[i] = 0;
		}
	}
}


// UndoRedo_Put: Allocates memory and saves the map structures
//----------------------------------------------------------------------------
int APIENTRY UndoRedo_Put(MAPTHING* things, int numthings, MAPLINEDEF* linedefs, int numlinedefs,
						  MAPSIDEDEF* sidedefs, int numsidedefs, MAPVERTEX* vertices, int numvertices,
						  MAPSECTOR* sectors, int numsectors)
{
	int index = -1;
	int s;
	
	// Find free index
	for(int i = 0; i < BACKUPLIMIT; i++)
	{
		// Check if free
		if(backup_things[i] == NULL)
		{
			// Use this index
			index = i;
			break;
		}
	}
	
	// Check if an index was found
	if(index > -1)
	{
		// Allocate memory
		backup_things[index] = new MAPTHING[numthings];
		backup_linedefs[index] = new MAPLINEDEF[numlinedefs];
		backup_sidedefs[index] = new MAPSIDEDEF[numsidedefs];
		backup_vertices[index] = new MAPVERTEX[numvertices];
		backup_sectors[index] = new MAPSECTOR[numsectors];
		
		// Copy memory
		memcpy(backup_things[index], things, sizeof(MAPTHING) * numthings);
		memcpy(backup_linedefs[index], linedefs, sizeof(MAPLINEDEF) * numlinedefs);
		memcpy(backup_sidedefs[index], sidedefs, sizeof(MAPSIDEDEF) * numsidedefs);
		memcpy(backup_vertices[index], vertices, sizeof(MAPVERTEX) * numvertices);
		memcpy(backup_sectors[index], sectors, sizeof(MAPSECTOR) * numsectors);
		
		// Keep number of items
		backup_numthings[index] = numthings;
		backup_numlinedefs[index] = numlinedefs;
		backup_numsidedefs[index] = numsidedefs;
		backup_numvertices[index] = numvertices;
		backup_numsectors[index] = numsectors;
		
		// Create new memory for each sidedef texture string
		for(s = 0; s < numsidedefs; s++)
		{
			if(sidedefs[s].lower != NULL)
			{
				backup_sidedefs[index][s].lower = new wchar_t[wcslen(sidedefs[s].lower) + 1];
				wcscpy(backup_sidedefs[index][s].lower, sidedefs[s].lower);
			}
			if(sidedefs[s].middle != NULL)
			{
				backup_sidedefs[index][s].middle = new wchar_t[wcslen(sidedefs[s].middle) + 1];
				wcscpy(backup_sidedefs[index][s].middle, sidedefs[s].middle);
			}
			if(sidedefs[s].upper != NULL)
			{
				backup_sidedefs[index][s].upper = new wchar_t[wcslen(sidedefs[s].upper) + 1];
				wcscpy(backup_sidedefs[index][s].upper, sidedefs[s].upper);
			}
		}
		
		// Create new memory for each sector flat string
		for(s = 0; s < numsectors; s++)
		{
			if(sectors[s].tceiling != NULL)
			{
				backup_sectors[index][s].tceiling = new wchar_t[wcslen(sectors[s].tceiling) + 1];
				wcscpy(backup_sectors[index][s].tceiling, sectors[s].tceiling);
			}
			if(sectors[s].tfloor != NULL)
			{
				backup_sectors[index][s].tfloor = new wchar_t[wcslen(sectors[s].tfloor) + 1];
				wcscpy(backup_sectors[index][s].tfloor, sectors[s].tfloor);
			}
		}
		
		// Done
		return index;
	}
	else
	{
		// Out of reserved memory!
		return -1;
	}
}


// UndoRedo_GetSizes: Reads the map structures
//----------------------------------------------------------------------------
void APIENTRY UndoRedo_GetSizes(int index, int &numthings, int &numlinedefs, int &numsidedefs,
								int &numvertices, int &numsectors)
{
	// Set number of items
	numthings = backup_numthings[index];
	numlinedefs = backup_numlinedefs[index];
	numsidedefs = backup_numsidedefs[index];
	numvertices = backup_numvertices[index];
	numsectors = backup_numsectors[index];
}


// UndoRedo_GetImages: Reads the map structures
//----------------------------------------------------------------------------
void APIENTRY UndoRedo_GetImages(int index, MAPTHING* things, MAPLINEDEF* linedefs,
								 MAPSIDEDEF* sidedefs, MAPVERTEX* vertices, MAPSECTOR* sectors)
{
	int s, ns;
	MAPSIDEDEF* osd;
	MAPSIDEDEF* bsd;
	MAPSECTOR* osc;
	MAPSECTOR* bsc;
	
	// Copy memory
	memcpy(things, backup_things[index], sizeof(MAPTHING) * backup_numthings[index]);
	memcpy(linedefs, backup_linedefs[index], sizeof(MAPLINEDEF) * backup_numlinedefs[index]);
	//memcpy(sidedefs, backup_sidedefs[index], sizeof(MAPSIDEDEF) * backup_numsidedefs[index]);
	memcpy(vertices, backup_vertices[index], sizeof(MAPVERTEX) * backup_numvertices[index]);
	//memcpy(sectors, backup_sectors[index], sizeof(MAPSECTOR) * backup_numsectors[index]);
	
	// Copy sidedefs (strings are a special case so we cant use copymemory here)
	ns = backup_numsidedefs[index];
	for(s = 0; s < ns; s++)
	{
		// References
		osd = &sidedefs[s];
		bsd = &backup_sidedefs[index][s];
		
		// Normal values
		osd->linedef = bsd->linedef;
		osd->sector = bsd->sector;
		osd->tx = bsd->tx;
		osd->ty = bsd->ty;
		
		// Strings
		if(bsd->lower != 0) wcsncpy(osd->lower, bsd->lower, wcslen(bsd->lower));
		if(bsd->middle != 0) wcsncpy(osd->middle, bsd->middle, wcslen(bsd->middle));
		if(bsd->upper != 0) wcsncpy(osd->upper, bsd->upper, wcslen(bsd->upper));
	}
	
	// Copy sectors (strings are a special case so we cant use copymemory here)
	ns = backup_numsectors[index];
	for(s = 0; s < ns; s++)
	{
		// References
		osc = &sectors[s];
		bsc = &backup_sectors[index][s];
		
		// Normal values
		osc->brightness = bsc->brightness;
		osc->hceiling = bsc->hceiling;
		osc->hfloor = bsc->hfloor;
		osc->selected = bsc->selected;
		osc->special = bsc->special;
		osc->tag = bsc->tag;
		
		// Strings
		if(bsc->tfloor != 0) wcsncpy(osc->tfloor, bsc->tfloor, wcslen(bsc->tfloor));
		if(bsc->tceiling != 0) wcsncpy(osc->tceiling, bsc->tceiling, wcslen(bsc->tceiling));
	}
}


// UndoRedo_Delete: Destroys (frees) a saved backup
//----------------------------------------------------------------------------
void APIENTRY UndoRedo_Delete(int index)
{
	int s, ns;
	
	// Release memory of sidedef strings
	ns = backup_numsidedefs[index];
	for(s = 0; s < ns; s++)
	{
		if(backup_sidedefs[index][s].lower != NULL) delete[] backup_sidedefs[index][s].lower;
		if(backup_sidedefs[index][s].middle != NULL) delete[] backup_sidedefs[index][s].middle;
		if(backup_sidedefs[index][s].upper != NULL) delete[] backup_sidedefs[index][s].upper;
	}
	
	// Release memory of sector strings
	ns = backup_numsectors[index];
	for(s = 0; s < ns; s++)
	{
		if(backup_sectors[index][s].tfloor != NULL) delete[] backup_sectors[index][s].tfloor;
		if(backup_sectors[index][s].tceiling != NULL) delete[] backup_sectors[index][s].tceiling;
	}
	
	// Release memory
	delete[] backup_things[index];
	delete[] backup_linedefs[index];
	delete[] backup_sidedefs[index];
	delete[] backup_vertices[index];
	delete[] backup_sectors[index];
	
	// Zero pointers and numbers
	backup_things[index] = NULL;
	backup_linedefs[index] = NULL;
	backup_sidedefs[index] = NULL;
	backup_vertices[index] = NULL;
	backup_sectors[index] = NULL;
	backup_numthings[index] = 0;
	backup_numlinedefs[index] = 0;
	backup_numsidedefs[index] = 0;
	backup_numvertices[index] = 0;
	backup_numsectors[index] = 0;
}