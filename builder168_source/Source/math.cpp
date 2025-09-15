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

// Headers
#include <windows.h>
#include <objbase.h>
#include "data.h"
#include "math.h"
#include "const.h"


// IntersectLineXZ: Same as IntersectLine, but uses X and Z cooridnates
// Returns -1 when not intersection, otherwise returns the distance in percent (0.0 - 1.0)
//-----------------------------------------------------------------------------
float __fastcall IntersectLineXZ(FPOINT3D &r1, FPOINT3D &r2, float &x3, float &z3, float &x4, float &z4)
{
	FPOINT3D r3, r4;
	r3.x = r1.x; r3.y = r1.z; r3.z = r1.y;
	r4.x = r2.x; r4.y = r2.z; r4.z = r2.y;
	return IntersectLine(r3, r4, x3, z3, x4, z4);
}


// IntersectLine: Tests a line with a ray.
// Returns -1 when not intersection, otherwise returns the distance in percent (0.0 - 1.0)
//-----------------------------------------------------------------------------
float __fastcall IntersectLine(FPOINT3D &r1, FPOINT3D &r2, float &x3, float &y3, float &x4, float &y4)
{
	// Calculate divider
	float div = (y4 - y3) * (r2.x - r1.x) - (x4 - x3) * (r2.y - r1.y);
	
	// Can this be tested?
	if(abs(div) > 0.00001f)
	{
		// Calculate the intersection distance from the ray
		float u_ray = ((x4 - x3) * (r1.y - y3) - (y4 - y3) * (r1.x - x3)) / div;
		
		// Calculate the intersection distance from the line
		float u_line = ((r2.x - r1.x) * (r1.y - y3) - (r2.y - r1.y) * (r1.x - x3)) / div;
		
		// Return the distance on the ray if intersecting, or return -1
		if((u_ray >= 0.0f) && (u_ray <= 1.0f) && (u_line >= 0.0f) && (u_line <= 1.0f)) return u_ray; else return -1.0f;
	}
	else
	{
		// Unable to detect intersection
		return -1.0f;
	}
}


// IntersectLineA: Tests a line with a ray.
// Returns -1 when not intersection, otherwise returns the distance in percent (0.0 - 1.0)
//-----------------------------------------------------------------------------
float APIENTRY IntersectLineA(float x1, float y1, float x2, float y2, float x3, float y3, float x4, float y4)
{
	// Calculate divider
	float div = (y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1);
	
	// Can this be tested?
	if(abs(div) > 0.00001f)
	{
		// Calculate the intersection distance from the ray
		float u_ray = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / div;
		
		// Calculate the intersection distance from the line
		float u_line = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / div;
		
		// Return the distance on the ray if intersecting, or return -1
		if((u_ray >= 0.0f) && (u_ray <= 1.0f) && (u_line >= 0.0f) && (u_line <= 1.0f)) return u_ray; else return -1.0f;
	}
	else
	{
		// Unable to detect intersection
		return -1.0f;
	}
}


// IntersectLine: Tests a line with a ray.
// Returns 1 when intersecting, 0 when not
//-----------------------------------------------------------------------------
int __fastcall IntersectLine(MAPVERTEX* v1, MAPVERTEX* v2, int x3, int y3, int x4, int y4)
{
	// Calculate divider
	float div = (y4 - y3) * (v2->x - v1->x) - (x4 - x3) * (v2->y - v1->y);
	
	// Can this be tested?
	if(abs(div) > 0.00001f)
	{
		// Calculate the intersection distance from the ray
		float u_ray = ((x4 - x3) * (v1->y - y3) - (y4 - y3) * (v1->x - x3)) / div;
		
		// Calculate the intersection distance from the line
		float u_line = ((v2->x - v1->x) * (v1->y - y3) - (v2->y - v1->y) * (v1->x - x3)) / div;
		
		// Return the 1 if intersecting, or return 0
		if((u_ray >= 0.0f) && (u_ray <= 1.0f) && (u_line >= 0.0f) && (u_line <= 1.0f)) return 1; else return 0;
	}
	else
	{
		// Unable to detect intersection
		return 0;
	}
}


// clipline: Clips line coordinates by the given dimensions and tests visibility.
// Returns True when the new line coordinates are visible, False when rejected.
//----------------------------------------------------------------------------
int clipline(float &x1, float &y1, float &x2, float &y2, float xmax, float ymax)
{
	// Clipping codes for each point
	int cc, cc1, cc2;
	float x, y;
	int result = false;
	
	// Determine clipping code for start point
	if(x1 < 0.0f) cc1 = CLIPLINE_LEFT; else if(x1 > xmax) cc1 = CLIPLINE_RIGHT; else cc1 = 0;
	if(y1 < 0.0f) cc1 |= CLIPLINE_TOP; else if(y1 > ymax) cc1 |= CLIPLINE_BOTTOM;
	
	// Determine clipping code for end point
	if(x2 < 0.0f) cc2 = CLIPLINE_LEFT; else if(x2 > xmax) cc2 = CLIPLINE_RIGHT; else cc2 = 0;
	if(y2 < 0.0f) cc2 |= CLIPLINE_TOP; else if(y2 > ymax) cc2 |= CLIPLINE_BOTTOM;
	
	// Return false if completely hidden
	if(cc1 & cc2) return false;
	
	// Continue until solved
	for(int i = 4; i > 0; i--)
	{
		// Return true if completely visible
		if((cc1 == 0) && (cc2 == 0))
		{
			// Return true
			return true;
		}
		else
		{
			// Get one of the points code that is outside the view
			if(cc1) cc = cc1; else cc = cc2;
			
			// Check what edge clips the line
			if(cc & CLIPLINE_TOP)
			{
				// Find intersection point
				x = x1 + (x2 - x1) * (0.0f - y1) / (y2 - y1);
				y = 0.0001f;
			}
			else if(cc & CLIPLINE_BOTTOM)
			{
				// Find intersection point
				x = x1 + (x2 - x1) * (ymax - y1) / (y2 - y1);
				y = ymax - 0.0001f;
			}
			else if(cc & CLIPLINE_RIGHT)
			{
				// Find intersection point
				y = y1 + (y2 - y1) * (xmax - x1) / (x2 - x1);
				x = xmax - 0.0001f;
			}
			else if(cc & CLIPLINE_LEFT)
			{
				// Find intersection point
				y = y1 + (y2 - y1) * (0.0f - x1) / (x2 - x1);
				x = 0.0001f;
			}
			
			// Check what point was used for the test
			if(cc == cc1)
			{
				// Cut the line at intersection
				x1 = x;
				y1 = y;
				
				// Recalculate clipping code
				if(x1 < 0.0f) cc1 = CLIPLINE_LEFT; else if(x1 > xmax) cc1 = CLIPLINE_RIGHT; else cc1 = 0;
				if(y1 < 0.0f) cc1 |= CLIPLINE_TOP; else if(y1 > ymax) cc1 |= CLIPLINE_BOTTOM;
			}
			else
			{
				// Cut the line at intersection
				x2 = x;
				y2 = y;
				
				// Recalculate clipping code
				if(x2 < 0.0f) cc2 = CLIPLINE_LEFT; else if(x2 > xmax) cc2 = CLIPLINE_RIGHT; else cc2 = 0;
				if(y2 < 0.0f) cc2 |= CLIPLINE_TOP; else if(y2 > ymax) cc2 |= CLIPLINE_BOTTOM;
			}
		}
	}
	
	// Somehow the line could not be cut to fit, so reject the line
	return false;
}


// pointinsidedefs: This tests if a point lies within or on the edge of
// a polygon of sidedefs
//----------------------------------------------------------------------------
bool pointinsidedefs(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int vertex, int* sideslist, int numsides)
{
	int inside = 0;
	MAPVERTEX p = vertices[vertex];
	
	// Go for all sidedefs
	for(int s = 0; s < numsides; s++)
	{
		// Get the line vertices
		int v1 = linedefs[sidedefs[sideslist[s]].linedef].v1;
		int v2 = linedefs[sidedefs[sideslist[s]].linedef].v2;
		
		// If the given vertex is one of the polygon vertices,
		// then the vertex is always considered on the edge of the polygon
		if((vertex == v1) || (vertex == v2)) return true;
		
		// Get coordinates
		float x1 = vertices[v1].x;
		float y1 = vertices[v1].y;
		float x2 = vertices[v2].x;
		float y2 = vertices[v2].y;
		
		// Check if the sidedefs crosses the point's Y axis
		if(((y1 < p.y) && (y2 >= p.y)) || ((y2 < p.y) && (y1 >= p.y)))
		{
			// Check if the sidedef is at the left of the point
			//if(x1 + (p.y - y1) / (y2 - y1) * (x2 - x1) < p.x) inside++;
			if(p.x <= (x2 - x1) * (p.y - y1) / (y2 - y1) + x1) inside++;
		}
	}
	
	// Return result
	return (inside & 1);
}

