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
#define MAX_CLIPPERS				300
#define MAX_SSECTORS_VERTICES			200
#define MAX_MAPBOUNDS				65536

// Includes
#include <windows.h>
#include <objbase.h>
#include "const.h"
#include "data.h"
#include "math.h"


// TriangulateSSector: Makes the polygon for the given subsector.
//-----------------------------------------------------------------------------
RECT __fastcall TriangulateSSector(POINT* vertices, MAPSEG* segs, MAPSSECTOR* ssector, MAPSPLIT* splitslist, int splits)
{
	int i, k, poly_numvertices, numclippers, numv;
	FPOINT* poly_vertices;
	CLIPPER* clippers;
	MAPSEG* sg;
	RECT bounds;
	int sidelist[MAX_CLIPPERS];
	int startIdx, endIdx, previdx;
	
	// This code is almost copied from the Doomsday engine.
	// I modified it for my editor and made a couple adjustments.
	
	// Calculate total clippers and reserve memory
	numclippers = splits + ssector->numsegs;
	clippers = (CLIPPER*)malloc(numclippers * sizeof(CLIPPER));
	
	// Limit the clippers
	if(numclippers > MAX_CLIPPERS) numclippers = MAX_CLIPPERS;
	
	// Create a list of all clipping lines
	// First the splits
	for(i = 0; i < splits; i++)
	{
		clippers[i].x = (float)splitslist[i].x;
		clippers[i].y = (float)splitslist[i].y;
		clippers[i].dx = (float)splitslist[i].dx;
		clippers[i].dy = (float)splitslist[i].dy;
	}
	
	// Then the segs
	for(i = 0; i < ssector->numsegs; i++)
	{
		sg = &segs[ssector->startseg + i];
		clippers[splits + i].x = (float)vertices[sg->v1].x;
		clippers[splits + i].y = (float)vertices[sg->v1].y;
		clippers[splits + i].dx = (float)vertices[sg->v2].x - (float)vertices[sg->v1].x;
		clippers[splits + i].dy = (float)vertices[sg->v2].y - (float)vertices[sg->v1].y;
	}
	
	// Create a simple polygon that covers the entire map
	poly_numvertices = 4;
	poly_vertices = (FPOINT*)malloc(poly_numvertices * sizeof(FPOINT));
	poly_vertices[0].x = -MAX_MAPBOUNDS;
	poly_vertices[0].y = MAX_MAPBOUNDS;
	poly_vertices[1].x = MAX_MAPBOUNDS;
	poly_vertices[1].y = MAX_MAPBOUNDS;
	poly_vertices[2].x = MAX_MAPBOUNDS;
	poly_vertices[2].y = -MAX_MAPBOUNDS;
	poly_vertices[3].x = -MAX_MAPBOUNDS;
	poly_vertices[3].y = -MAX_MAPBOUNDS;
	
	
	// Copy the number of vertices
	numv = poly_numvertices;
	
	// We'll clip the polygon with each of the divlines. The left side of
	// each divline is discarded.
	for(i = 0; i < numclippers; i++)
	{
		// Get the clipper
		CLIPPER* curclip = &clippers[i];
		
		// First we'll determine the side of each vertex.
		// Set to 1 for right or ON the line, 0 for left of the line.
		for(k = 0; k < numv; k++) sidelist[k] = ((curclip->y - poly_vertices[k].y) * curclip->dx - (curclip->x - poly_vertices[k].x) * curclip->dy >= 0);
		
		// Go for all vertices
		for(k = 0; k < numv; k++)
		{
			// Determine start and end
			startIdx = k;
			endIdx = k + 1;
			if(endIdx == numv) endIdx = 0;	// Wrap-around
			
			// Clipping will happen when the ends are on different sides.
			if(sidelist[startIdx] != sidelist[endIdx])
			{
				FPOINT newvert;
				
				// Find the intersection point of intersecting lines
				intercept_vertex(&poly_vertices[startIdx], &poly_vertices[endIdx], curclip, &newvert);
				
				// Add the new vertex
				poly_vertices = (FPOINT*)realloc(poly_vertices, (++numv) * sizeof(FPOINT));
				memmove(&poly_vertices[endIdx + 1], &poly_vertices[endIdx], (numv - endIdx - 1) * sizeof(FPOINT));
				memcpy(&poly_vertices[endIdx], &newvert, sizeof(FPOINT));
				
				// Modify the sides list
				memmove(&sidelist[endIdx + 1], &sidelist[endIdx], (numv - endIdx - 1) * sizeof(int));
				sidelist[endIdx] = 1;
				
				// Skip over the new vertex
				k++;
			}
		}
		
		// Now we must discard the points that are on the wrong side
		for(k = 0; k < numv; k++)
		{
			if(!sidelist[k])
			{
				memmove(&poly_vertices[k], &poly_vertices[k + 1], (numv - k - 1) * sizeof(FPOINT));
				memmove(&sidelist[k], &sidelist[k + 1], (numv - k - 1) * sizeof(int));
				numv--;
				k--;
			}
		}
	}
	
	// Let the boundaries begin with the first vertex
	bounds.left = (int)poly_vertices[0].x;
	bounds.right = (int)poly_vertices[0].x;
	bounds.top = (int)poly_vertices[0].y;
	bounds.bottom = (int)poly_vertices[0].y;
	
	// Screen out consecutive identical points
	for(i = 0; i < numv; i++)
	{
		previdx = i - 1;
		if(previdx < 0) previdx = numv - 1;
		if((fabsf(poly_vertices[i].x - poly_vertices[previdx].x) < 0.01) &&
		   (fabsf(poly_vertices[i].y - poly_vertices[previdx].y) < 0.01))
		{
			// This point must be removed
			memmove(&poly_vertices[i], &poly_vertices[i + 1], (numv - i - 1) * sizeof(FPOINT));
			numv--;
			i--;
		}
		else
		{
			// Apply to boundaries
			if((float)bounds.left > poly_vertices[i].x) bounds.left = (int)floor(poly_vertices[i].x);
			if((float)bounds.right < poly_vertices[i].x) bounds.right = (int)ceil(poly_vertices[i].x);
			if((float)bounds.top > poly_vertices[i].y) bounds.top = (int)floor(poly_vertices[i].y);
			if((float)bounds.bottom < poly_vertices[i].y) bounds.bottom = (int)ceil(poly_vertices[i].y);
		}
	}
	
	// Set and limit the new number of vertices
	poly_numvertices = numv;
	if(poly_numvertices > MAX_SSECTORS_VERTICES) poly_numvertices = MAX_SSECTORS_VERTICES;
	
	// Now simply copy the polygon on the subsector
	ssector->numvertices = poly_numvertices;
	memcpy(ssector->vertices, poly_vertices, poly_numvertices * sizeof(FPOINT));
	
	// Clean up
	free(clippers);
	free(poly_vertices);
	
	// Return the bounds
	return bounds;
}
