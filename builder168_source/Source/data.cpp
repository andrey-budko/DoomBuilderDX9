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
#include <limits.h>
#include <string>
#include <hash_map>
#include "const.h"
#include "data.h"
#include "math.h"


// This reallocates an int array and preserves any contents
//----------------------------------------------------------------------------
void __fastcall ReallocIntP(int* &intarray, int oldlength, int newlength)
{
	// Allocate new memory
	int* newarray = new int[newlength];
	
	// Check if an old list is given
	if(intarray != NULL)
	{
		// Copy contents
		memcpy(newarray, intarray, oldlength * sizeof(int));
		
		// Kill the old list
		delete[] intarray;
	}
	
	// Apply the new list
	intarray = newarray;
}


// ThingFiltered: Returns 1 if a given thing must be shown
//----------------------------------------------------------------------------
int __fastcall ThingFiltered(MAPTHING* thing, int filterthings, THINGFILTERS* filter)
{
	// Check if filtering
	if(filterthings)
	{
		// Test category
		if( (filter->category == -1) || (thing->category == filter->category) )
		{
			// Check filter mode
			switch(filter->filtermode)
			{
				case 0: return ((thing->flag & filter->flags) != 0) || (thing->flag == 0);
				case 1: return ((thing->flag & filter->flags) == filter->flags);
				case 2: return (thing->flag == filter->flags);
				default: return 0;
			}
		}
		else
		{
			// Wrong category
			return 0;
		}
	}
	else
	{
		// Always shown
		return 1;
	}
}


// RoundVertices: Rounds the vertices to the nearest whole number
//-----------------------------------------------------------------------------
void APIENTRY RoundVertices(MAPVERTEX* vertices, int numvertices)
{
	// Go for all vertices
	for(int v = 0; v < numvertices; v++)
	{
		// Round vertex
		vertices[v].x = (float)(int)(vertices[v].x);
		vertices[v].y = (float)(int)(vertices[v].y);
	}
}


// Rereference_Vertices: Changes specified references for Vertices on Linedefs
//-----------------------------------------------------------------------------
void APIENTRY Rereference_Vertices(MAPLINEDEF* linedefs, int numlinedefs, int oldref, int newref)
{
	// Go for all linedefs
	for(int ld = 0; ld < numlinedefs; ld++)
	{
		// Rereference if S1 is old reference
		if(linedefs[ld].v1 == oldref) linedefs[ld].v1 = newref;
		if(linedefs[ld].v2 == oldref) linedefs[ld].v2 = newref;
	}
}


// Rereference_Sidedefs: Changes specified references for Sidedefs on Linedefs
//-----------------------------------------------------------------------------
void APIENTRY Rereference_Sidedefs(MAPLINEDEF* linedefs, int numlinedefs, int oldref, int newref)
{
	// Go for all linedefs
	for(int ld = 0; ld < numlinedefs; ld++)
	{
		// Rereference if S1 is old reference
		if(linedefs[ld].s1 == oldref) linedefs[ld].s1 = newref;
		if(linedefs[ld].s2 == oldref) linedefs[ld].s2 = newref;
	}
}


// Rereference_SidedefsLinedef: Changes specified references for Linedefs on Sidedefs
//-----------------------------------------------------------------------------
void APIENTRY Rereference_SidedefsLinedef(MAPSIDEDEF* sidedefs, int numsidedefs, int oldref, int newref)
{
	// Go for all sidedefs
	for(int sd = 0; sd < numsidedefs; sd++)
	{
		// Rereference if linedef is old reference
		if(sidedefs[sd].linedef == oldref) sidedefs[sd].linedef = newref;
	}
}


// Rereference_Sectors: Changes specified references for Sectors on Sidedefs
//-----------------------------------------------------------------------------
void APIENTRY Rereference_Sectors(MAPSIDEDEF* sidedefs, int numsidedefs, int oldref, int newref)
{
	// Go for all sidedefs
	for(int sd = 0; sd < numsidedefs; sd++)
	{
		// Rereference if sector is old reference
		if(sidedefs[sd].sector == oldref) sidedefs[sd].sector = newref;
	}
}


// CountSectorSidedefs: Returns the number of sidedefs referring to a given sector
//-----------------------------------------------------------------------------
int APIENTRY CountSectorSidedefs(MAPSIDEDEF* sidedefs, int numsidedefs, int sector)
{
	int count = 0;
	
	// Count all sidedefs that refer to the givens ector
	for(int sd = 0; sd < numsidedefs; sd++) { if(sidedefs[sd].sector == sector) count++; }
	
	// Return result
	return count;
}


// CountVertexLinedefs: Returns the number of linedefs referring to a given vertex
//-----------------------------------------------------------------------------
int APIENTRY CountVertexLinedefs(MAPLINEDEF* linedefs, int numlinedefs, int vertex)
{
	int count = 0;
	
	// Count all lines that refer to the given vertex
	for(int ld = 0; ld < numlinedefs; ld++)	{ if((linedefs[ld].v1 == vertex) || (linedefs[ld].v2 == vertex)) count++; }
	
	// Return result
	return count;
}


// ResetSelections: Sets all selected properties to 0
//-----------------------------------------------------------------------------
void APIENTRY ResetSelections(MAPTHING* things, int numthings, MAPLINEDEF* linedefs, int numlinedefs,
							  MAPVERTEX* vertices, int numvertices, MAPSECTOR* sectors, int numsectors)
{
	int i;
	
	// Reset all selected properties
	if(numthings) for(i = 0; i < numthings; i++) things[i].selected = 0;
	if(numlinedefs) for(i = 0; i < numlinedefs; i++) linedefs[i].selected = 0;
	if(numvertices) for(i = 0; i < numvertices; i++) vertices[i].selected = 0;
	if(numsectors) for(i = 0; i < numsectors; i++) sectors[i].selected = 0;
}


// NearestVertex: Returns the nearest vertex index
//----------------------------------------------------------------------------
int APIENTRY NearestVertex(int x, int y, MAPVERTEX* vertices, int numvertices, int &dist)
{
	int foundvertex = -1;
	int founddistance = INT_MAX;
	int d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all vertices
	for(int v = 0; v < numvertices; v++)
	{
		// Calculate distance
		d = (int)distance(int(vertices[v].x), int(vertices[v].y), x, y);
		
		// Check if closer
		if(d < founddistance)
		{
			// Found a closer match
			foundvertex = v;
			founddistance = d;
		}
	}
	
	// Return result
	dist = founddistance;
	return foundvertex;
}


// NearestOtherVertex: Returns the vertex index nearest to given vertex index
//----------------------------------------------------------------------------
int APIENTRY NearestOtherVertex(int vx, MAPVERTEX* vertices, int numvertices, int &dist)
{
	int foundvertex = -1;
	int founddistance = INT_MAX;
	int d, v;
	int x = (int)vertices[vx].x;
	int y = (int)vertices[vx].y;
	
	// Go for all vertices before vx
	for(v = 0; v < vx; v++)
	{
		// Calculate distance
		d = (int)distance(int(vertices[v].x), int(vertices[v].y), x, y);
		
		// Check if closer
		if(d < founddistance)
		{
			// Found a closer match
			foundvertex = v;
			founddistance = d;
		}
	}
	
	// Go for all vertices after vx
	for(v = vx + 1; v < numvertices; v++)
	{
		// Calculate distance
		d = (int)distance(int(vertices[v].x), int(vertices[v].y), x, y);
		
		// Check if closer
		if(d < founddistance)
		{
			// Found a closer match
			foundvertex = v;
			founddistance = d;
		}
	}
	
	// Return result
	dist = founddistance;
	return foundvertex;
}


// NearestSelectedVertex: Returns the nearest selected vertex index
//----------------------------------------------------------------------------
int APIENTRY NearestSelectedVertex(int x, int y, MAPVERTEX* vertices, int numvertices, int &dist)
{
	int foundvertex = -1;
	int founddistance = INT_MAX;
	int d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all vertices
	for(int v = 0; v < numvertices; v++)
	{
		// Check if selected
		if(vertices[v].selected)
		{
			// Calculate distance
			d = (int)distance(int(vertices[v].x), int(vertices[v].y), x, y);
			
			// Check if closer
			if(d < founddistance)
			{
				// Found a closer match
				foundvertex = v;
				founddistance = d;
			}
		}
	}
	
	// Return result
	dist = founddistance;
	return foundvertex;
}


// NearestUnselectedVertex: Returns the nearest unselected vertex index
//----------------------------------------------------------------------------
int APIENTRY NearestUnselectedVertex(int x, int y, MAPVERTEX* vertices, int numvertices, int &dist)
{
	int foundvertex = -1;
	int founddistance = INT_MAX;
	float fx = (float)x, fy = (float)y;
	int d;
	
	// Correct the Y axis
	fy = -fy;
	
	// Go for all vertices
	for(int v = 0; v < numvertices; v++)
	{
		// Check if not selected
		if(vertices[v].selected == 0)
		{
			// Calculate distance
			d = (int)distance(vertices[v].x, vertices[v].y, fx, fy);
			
			// Check if closer
			if(d < founddistance)
			{
				// Found a closer match
				foundvertex = v;
				founddistance = d;
			}
		}
	}
	
	// Return result
	dist = founddistance;
	return foundvertex;
}


// OverlappingUnselectedVertex: Returns the overlapping unselected vertex index
//----------------------------------------------------------------------------
int APIENTRY OverlappingUnselectedVertex(MAPVERTEX* vertices, int numvertices, int tv)
{
	float fx = vertices[tv].x;
	float fy = vertices[tv].y;
	
	// Go for all vertices
	for(int v = 0; v < numvertices; v++)
	{
		// Check if not selected
		if(vertices[v].selected == 0)
		{
			// Check if overlapping
			if((vertices[v].x == fx) && (vertices[v].y == fy))
			{
				// if this is not the same vertex as the source,
				// then this is an overlapping vertex
				if(v != tv) return v;
			}
		}
	}
	
	// Return -1 (nothing found)
	return -1;
}


// OverlappingUnselectedLinedef: Returns the overlapping unselected linedef index
//----------------------------------------------------------------------------
int APIENTRY OverlappingUnselectedLinedef(MAPLINEDEF* linedefs, int numlinedefs, int tl)
{
	float tlv1 = (float)linedefs[tl].v1;
	float tlv2 = (float)linedefs[tl].v2;
	
	// Go for all linedefs
	for(int ld = 0; ld < numlinedefs; ld++)
	{
		// Check if not selected
		if(linedefs[ld].selected == 0)
		{
			// Check if overlapping
			if(((linedefs[ld].v1 == tlv1) && (linedefs[ld].v2 == tlv2)) ||
			   ((linedefs[ld].v1 == tlv2) && (linedefs[ld].v2 == tlv1)))
			{
				// if this is not the same linedef as the source,
				// then this is an overlapping linedef
				if(ld != tl) return ld;
			}
		}
	}
	
	// Return -1 (nothing found)
	return -1;
}


// NearestThing: Returns the nearest thing index
//----------------------------------------------------------------------------
int APIENTRY NearestThing(int x, int y, MAPTHING* things, int numthings, int &dist, int filterthings, THINGFILTERS* filter)
{
	int foundthing = -1;
	int founddistance = INT_MAX;
	int d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Check if any of the thing flags match any of the filter flags
		if(ThingFiltered(&things[th], filterthings, filter) || (things[th].selected != 0))
		{
			// Calculate distance
			d = (int)distance(things[th].x, things[th].y, x, y);
			
			// Check if closer
			if(d < founddistance)
			{
				// Found a closer match
				foundthing = th;
				founddistance = d;
			}
		}
	}
	
	// Return result
	dist = founddistance;
	return foundthing;
}


// NearestSelectedThing: Returns the nearest selected thing index
//----------------------------------------------------------------------------
int APIENTRY NearestSelectedThing(int x, int y, MAPTHING* things, int numthings, int &dist)
{
	int foundthing = -1;
	int founddistance = INT_MAX;
	int d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all vertices
	for(int th = 0; th < numthings; th++)
	{
		// Check if selected or all things are allowed
		if(things[th].selected)
		{
			// Calculate distance
			d = (int)distance(things[th].x, things[th].y, x, y);
			
			// Check if closer
			if(d < founddistance)
			{
				// Found a closer match
				foundthing = th;
				founddistance = d;
			}
		}
	}
	
	// Return result
	dist = founddistance;
	return foundthing;
}


// NearestUnselectedThing: Returns the nearest unselected thing index
//----------------------------------------------------------------------------
int APIENTRY NearestUnselectedThing(int x, int y, MAPTHING* things, int numthings, int &dist, int filterthings, THINGFILTERS* filter)
{
	int foundthing = -1;
	int founddistance = INT_MAX;
	int d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all vertices
	for(int th = 0; th < numthings; th++)
	{
		// Check if selected or all things are allowed
		if(things[th].selected == 0)
		{
			// Check if any of the thing flags match any of the filter flags
			if(ThingFiltered(&things[th], filterthings, filter))
			{
				// Calculate distance
				d = (int)distance(things[th].x, things[th].y, x, y);
				
				// Check if closer
				if(d < founddistance)
				{
					// Found a closer match
					foundthing = th;
					founddistance = d;
				}
			}
		}
	}
	
	// Return result
	dist = founddistance;
	return foundthing;
}


// NearestLinedef: Returns the nearest linedef index
//----------------------------------------------------------------------------
int APIENTRY NearestLinedef(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int &dist) { return NearestLinedefF(x, y, vertices, linedefs, numlinedefs, dist); }
int __fastcall NearestLinedefF(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int &dist)
{
	int foundlinedef = -1;
	float founddistance = 1000000;
	float d;
	
	// Correct the Y axis
	y = -y;
	
	// Go for all linedefs
	for(int l = 0; l < numlinedefs; l++)
	{
		// Get linedef vertices
		MAPVERTEX* v1 = &vertices[linedefs[l].v1];
		MAPVERTEX* v2 = &vertices[linedefs[l].v2];
		
		// Get shortest distance to linedef
		d = distance_to_line(v1->x, v1->y, v2->x, v2->y, (float)x, (float)y);
		
		// Check if closer but 'within' range
		if(d < founddistance)
		{
			// Found a closer match
			foundlinedef = l;
			founddistance = d;
		}
	}
	
	// Return result
	dist = (int)founddistance;
	return foundlinedef;
}


// NearestSelectedLinedef: Returns the nearest selected linedef index
//----------------------------------------------------------------------------
int APIENTRY NearestSelectedLinedef(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int* selected, int numselectedlinedefs, int &dist, int maxdist)
{
	int foundlinedef = -1;
	float founddistance = 1000000;
	float d;
	float fx = (float)x, fy = (float)y;
	int l;
	
	// Correct the Y axis
	fy = -fy;
	
	// Go for all linedefs
	for(int s = 0; s < numselectedlinedefs; s++)
	{
		// Get the linedef
		l = selected[s];
		
		// Get linedef vertices
		MAPVERTEX* v1 = &vertices[linedefs[l].v1];
		MAPVERTEX* v2 = &vertices[linedefs[l].v2];
		
		// Check if point is near the line
		if(point_near_line(v1->x, v1->y, v2->x, v2->y, fx, fy, maxdist))
		{
			// Get shortest distance to linedef
			d = distance_to_line(v1->x, v1->y, v2->x, v2->y, fx, fy);
			
			// Check if closer but 'within' range
			if(d < founddistance)
			{
				// Found a closer match
				foundlinedef = l;
				founddistance = d;
			}
		}
	}
	
	// Return result
	dist = (int)founddistance;
	return foundlinedef;
}


// NearestUnselectedLinedef: Returns the nearest unselected linedef index
//----------------------------------------------------------------------------
int APIENTRY NearestUnselectedLinedef(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int &dist, int maxdist)
{
	int foundlinedef = -1;
	float founddistance = 1000000;
	float fx = (float)x, fy = (float)y;
	float d;
	
	// Correct the Y axis
	fy = -fy;
	
	// Go for all linedefs
	for(int l = 0; l < numlinedefs; l++)
	{
		// Check if not selected
		if(linedefs[l].selected == 0)
		{
			// Get linedef vertices
			MAPVERTEX* v1 = &vertices[linedefs[l].v1];
			MAPVERTEX* v2 = &vertices[linedefs[l].v2];
			
			// Check if point is near the line
			if(point_near_line(v1->x, v1->y, v2->x, v2->y, fx, fy, maxdist))
			{
				// Get shortest distance to linedef
				d = distance_to_line(v1->x, v1->y, v2->x, v2->y, fx, fy);
				
				// Check if closer but 'within' range
				if(d < founddistance)
				{
					// Found a closer match
					foundlinedef = l;
					founddistance = d;
				}
			}
		}
	}
	
	// Return result
	dist = (int)founddistance;
	return foundlinedef;
}


// NearestUnselectedUnreferencedLinedef: Returns the nearest unselected and unreferenced linedef index
//----------------------------------------------------------------------------
int APIENTRY NearestUnselectedUnreferencedLinedef(int v, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int &dist, int maxdist)
{
	int foundlinedef = -1;
	float founddistance = 1000000;
	float d;
	
	// Make coordinates
	float y = vertices[v].y;
	float x = vertices[v].x;
	
	// Go for all linedefs
	for(int l = 0; l < numlinedefs; l++)
	{
		// Check if not selected
		if(linedefs[l].selected == 0)
		{
			// Check if vertices do not refer to this vertex
			if((linedefs[l].v1 != v) && (linedefs[l].v2 != v))
			{
				// Get linedef vertices
				MAPVERTEX* v1 = &vertices[linedefs[l].v1];
				MAPVERTEX* v2 = &vertices[linedefs[l].v2];
				
				// Check if near the line
				if(point_near_line(v1->x, v1->y, v2->x, v2->y, x, y, maxdist))
				{
					// Get shortest distance to linedef
					d = distance_to_line(v1->x, v1->y, v2->x, v2->y, x, y);
					
					// Check if closer but 'within' range
					if(d < founddistance)
					{
						// Found a closer match
						foundlinedef = l;
						founddistance = d;
					}
				}
			}
		}
	}
	
	// Return result
	dist = (int)founddistance;
	return foundlinedef;
}


// IntersectSector: Returns the intersecting sector index
//----------------------------------------------------------------------------
int APIENTRY IntersectSector(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs, int unselectedonly)
{
	int nld;
	
	// Check if only checking with unselected lines
	if(unselectedonly)
	{
		// Lets get the nearest unselected linedef
		int lddist = 0;
		nld = NearestUnselectedLinedef(x, y, vertices, linedefs, numlinedefs, lddist, ENDLESS_DISTANCE);
	}
	else
	{
		// Lets get the nearest linedef
		int lddist = 0;
		nld = NearestLinedefF(x, y, vertices, linedefs, numlinedefs, lddist);
	}
	
	// Only continue if any linedefs found
	if(nld > -1)
	{
		// Correct the Y axis
		y = -y;
		
		// Get the vertices
		MAPVERTEX* v1 = &vertices[linedefs[nld].v1];
		MAPVERTEX* v2 = &vertices[linedefs[nld].v2];
		
		// Check the side of the line
		if(side_of_line((int)v1->x, (int)v1->y, (int)v2->x, (int)v2->y, x, y) < 0)
		{
			// Front side (sidedef 1)
			// Return the referenced sector number if a sidedef is referenced
			if(linedefs[nld].s1 > -1) return sidedefs[linedefs[nld].s1].sector; else return -1;
		}
		else
		{
			// Back side (sidedef 2)
			// Return the referenced sector number if a sidedef is referenced
			if(linedefs[nld].s2 > -1) return sidedefs[linedefs[nld].s2].sector; else return -1;
		}
	}
	else
	{
		// Nothing intersected
		return -1;
	}
}


// LinedefBetweenVertices: Returns the first next linedefs that is between the given vertices
//----------------------------------------------------------------------------
int APIENTRY LinedefBetweenVertices(MAPLINEDEF* linedefs, int numlinedefs, int startlinedef, int vertex1, int vertex2, int excludeline)
{
	// Go for all linedefs
	for(int ld = startlinedef; ld < numlinedefs; ld++)
	{
		// Check linedef vertices
		if(((linedefs[ld].v1 == vertex1) && (linedefs[ld].v2 == vertex2)) ||
		   ((linedefs[ld].v2 == vertex1) && (linedefs[ld].v1 == vertex2))) if(ld != excludeline) return ld;
	}
	
	// None found
	return -1;
}


// SetAllThingSectors: Sets the sector property on all Things
//----------------------------------------------------------------------------
void APIENTRY SetAllThingSectors(MAPTHING* things, int numthings, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, MAPSIDEDEF* sidedefs)
{
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Set the thing sector
		things[th].sector = IntersectSector(things[th].x, -things[th].y, vertices, linedefs, sidedefs, numlinedefs, 0);
	}
}


// FurthestSectorVertex: Finds the furthest vertex from a line in sector
//----------------------------------------------------------------------------
int APIENTRY FurthestSectorVertex(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs, int line, int sector)
{ return FurthestSectorVertexF(vertices, linedefs, sidedefs, numlinedefs, line, sector); }
int FurthestSectorVertexF(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs, int line, int sector)
{
	int furthestvertex = -1;
	float furthestdistance = -1;
	float dx, dy, vx, vy, d;
	MAPVERTEX* v;
	int s1, s2;
	
	// Calculate line slope and coordinates
	vx = vertices[linedefs[line].v1].x;
	vy = vertices[linedefs[line].v1].y;
	dx = vx - vertices[linedefs[line].v2].x;
	dy = vy - vertices[linedefs[line].v2].y;
	
	// Go for all linedefs
	for(int ld = 0; ld < numlinedefs; ld++)
	{
		// Check if this line is with the given sector
		if(linedefs[ld].s1 > -1) s1 = sidedefs[linedefs[ld].s1].sector; else s1 = -1;
		if(linedefs[ld].s2 > -1) s2 = sidedefs[linedefs[ld].s2].sector; else s2 = -1;
		if((s1 == sector) || (s2 == sector))
		{
			// Test first vertex
			v = vertices + linedefs[ld].v1;
			d = abs(((vy - v->y) * dx) - ((vx - v->x) * dy));
			if(d > furthestdistance)
			{
				furthestdistance = d;
				furthestvertex = linedefs[ld].v1;
			}
			
			// Test second vertex
			v = vertices + linedefs[ld].v2;
			d = abs(((vy - v->y) * dx) - ((vx - v->x) * dy));
			if(d > furthestdistance)
			{
				furthestdistance = d;
				furthestvertex = linedefs[ld].v2;
			}
		}
	}
	
	// Return result
	return furthestvertex;
}




