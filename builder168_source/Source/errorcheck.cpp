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
#include "const.h"
#include "data.h"
#include "math.h"

// turn off some stupid warnings
#pragma warning (disable: 4244)					// conversion from 'type1' to 'type2', possible loss of data


// This is a list of all linedef numbers of a sector
int* sectorlines = NULL;	// linedef numbers
int* sectorsides = NULL;	// side of line attached  1=front 2=back
int sectorlinescap = 0;
int numsectorlines = 0;

// This is used to find the closest vertex to a sector leak
int lastvertex = -1;
int lastlength = 0;

// This is a list of found lines of a sector
// They will be used to test if lines of the unclosed
// parts are completely within the sector
int* looplines = NULL;
int numlooplines = 0;


// This test a square area against all impossible lines
// Returns the colliding linedef number or -1 when not colliding
//----------------------------------------------------------------------------
int APIENTRY TestStuckedThing(MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlines, int x, int y, int radius)
{
	MAPVERTEX* v1;
	MAPVERTEX* v2;
	
	// Setup square coordinates
	int l = x - radius;
	int r = x + radius;
	int t = y - radius;
	int b = y + radius;
	
	// Go for all lines
	for(int ld = 0; ld < numlines; ld++)
	{
		// Test if the line is impassable
		if(linedefs[ld].flags & LDF_IMPASSIBLE)
		{
			// Get line coordinates
			v1 = &vertices[linedefs[ld].v1];
			v2 = &vertices[linedefs[ld].v2];
			
			// Test if the line is a potential collision
			if(!(((v1->x < l) && (v2->x < l)) ||
			     ((v1->x > r) && (v2->x > r)) ||
			     ((v1->y < t) && (v2->y < t)) ||
			     ((v1->y > b) && (v2->y > b))))
			{
				// When an endpoint of the line lies within the square
				// it is always a collision
				if(point_in_rect(v1->x, v1->y, l, r, t, b) ||
				   point_in_rect(v2->x, v2->y, l, r, t, b))
				{
					// This square collides with this line
					return ld;
				}
				else
				{
					// Perform 4 line vs line intersection tests
					if(IntersectLine(v1, v2, l, t, r, t) ||
					   IntersectLine(v1, v2, r, t, r, b) ||
					   IntersectLine(v1, v2, r, b, l, b) ||
					   IntersectLine(v1, v2, l, b, l, t))
					{
						// This square collides with this line
						return ld;
					}
				}
			}
		}
	}
	
	// No collision
	return -1;
}


// This creates the list of sector linedefs and sides
//----------------------------------------------------------------------------
void FindSectorSideLines(MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs, int sector)
{
	// Go for all linedefs
	for(int ld = 0; ld < numlinedefs; ld++)
	{
		// Check if the line has a front side
		if(linedefs[ld].s1 != -1)
		{
			// Check if at this sector
			if(sidedefs[linedefs[ld].s1].sector == sector)
			{
				// Check if reallocation is needed
				if(numsectorlines == sectorlinescap)
				{
					// Reallocate arrays
					ReallocIntP(sectorlines, sectorlinescap, sectorlinescap + 10);
					ReallocIntP(sectorsides, sectorlinescap, sectorlinescap + 10);
					sectorlinescap += 10;
				}
				
				// Add this line to the list
				sectorlines[numsectorlines] = ld;
				sectorsides[numsectorlines] = 1;
				numsectorlines++;
			}
		}
		
		// Check if the line has a back side
		if(linedefs[ld].s2 != -1)
		{
			// Check if at this sector
			if(sidedefs[linedefs[ld].s2].sector == sector)
			{
				// Check if reallocation is needed
				if(numsectorlines == sectorlinescap)
				{
					// Reallocate arrays
					ReallocIntP(sectorlines, sectorlinescap, sectorlinescap + 10);
					ReallocIntP(sectorsides, sectorlinescap, sectorlinescap + 10);
					sectorlinescap += 10;
				}
				
				// Add this line to the list
				sectorlines[numsectorlines] = ld;
				sectorsides[numsectorlines] = 2;
				numsectorlines++;
			}
		}
	}
}


// This is a recursive routine to find the loop
// Returns a pointer to the lines array when a loop
// was found and the recursion must be terminated.
// Used by FindLinedefLoop.
//----------------------------------------------------------------------------
int* NextVertexToLoop(const MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs,
					  int startvertex, int endvertex, int prevline, int prevside, int prevforward,
					  int sector, int* sideslist, int &numsideslist, int &lastvertex, int &longestlist)
{
	int ld, ev, i, s, sl, sd;
	int curforward, curside;
	int lastvertexcandidate = 1;
	
	// Go for all linedefs in sector
	for(sl = 0; sl < numsectorlines; sl++)
	{
		// Get linedef index and side
		ld = sectorlines[sl];
		sd = sectorsides[sl];
		
		// Check if this linedef is starting at this vertex
		if(linedefs[ld].v1 == startvertex)
		{
			// Previously going backward?
			if(prevforward == 0)
			{
				// Then we must switch sides
				if(prevside) curside = 0; else curside = 1;
			}
			else
			{
				// Same side
				curside = prevside;
			}
			
			// Now going forward
			curforward = 1;
		}
		// Check if this linedef is ending at this vertex
		else if(linedefs[ld].v2 == startvertex)
		{
			// Previously going forward?
			if(prevforward == 1)
			{
				// Then we must switch sides
				if(prevside) curside = 0; else curside = 1;
			}
			else
			{
				// Same side
				curside = prevside;
			}
			
			// Now going backward
			curforward = 0;
		}
		else
		{
			// Next line!
			continue;
		}
		
		// Skip if this is the same as previous line
		if(ld == prevline) continue;
		
		// Check if the side matches this line side
		if(curside + 1 == sd)
		{
			// Check what side to move along
			if(curside == 0)
			{
				// If already selected, then dont use it
				if((linedefs[ld].selected & 1) == 1) continue;
				
				// Get the sidedef index to move along
				s = linedefs[ld].s1;
			}
			else
			{
				// If already selected, then dont use it
				if((linedefs[ld].selected & 2) == 2) continue;
				
				// Get the sidedef index to move along
				s = linedefs[ld].s2;
			}
			
			// Other lines on this sector detected at this vertex,
			// so this means this vertex cant be a candidate for last vertex
			lastvertexcandidate = 0;
			
			// Go for all sidedefs already checked
			for(i = 0; i < numsideslist; i++)
			{
				// Check if this sidedef is already checked
				if(s == sideslist[i])
				{
					// Mark as invalid so it will be skipped
					s = -1;
					break;
				}
			}
			
			// If the sidedef is not marked as invalid...
			if(s != -1)
			{
				// This sidedef is part of our loop!
				
				// Create new array and add this line to it
				int newnumsides = numsideslist + 1;
				if(newnumsides > longestlist) longestlist = newnumsides;
				int* newlist = new int[newnumsides];
				if(sideslist != NULL) memcpy(newlist, sideslist, numsideslist * sizeof(int));
				newlist[numsideslist] = s;
				
				// Get the ending vertex after this line
				if(curforward) ev = linedefs[ld].v2; else ev = linedefs[ld].v1;
				
				// Check if the ending vertex is our destination
				if(ev == endvertex)
				{
					// Return the pointer to the array of complete lines
					numsideslist = newnumsides;
					return newlist;
				}
				else
				{
					// Continue tracking recursively along this line
					int* returnlist = NextVertexToLoop(linedefs, sidedefs, ev,
													   endvertex, ld, curside, curforward,
													   sector, newlist, newnumsides, lastvertex,
													   longestlist);
					
					// Clean up our list
					delete[] newlist;
					
					// When a result has been returned, terminate and return the result
					if(returnlist != NULL)
					{
						// Return result list and length of list
						numsideslist = newnumsides;
						return returnlist;
					}
				}
			}
		}
	}
	
	// Candidate for last vertex?
	if(lastvertexcandidate) lastvertex = startvertex;
	
	// Return nothing
	return NULL;
}


// FindLinedefLoop: Attempts to find a loop of linedefs for a given sector and start vertex
// Returns the list of lines when a valid loop was found, returns NULL when no valid loop found
// This also SELECTS the lines of the found loop!
//----------------------------------------------------------------------------
int* FindLinedefLoop(MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int startline,
					 int startside, int &lastvertex, int &longestlist, int &numlines)
{
	int startvertex, endvertex, startforward, sector;
	int i, ld;
	
	// Create initial list
	int* firstlist = new int[1];
	
	// Determine begin settings
	if(startside == 0)
	{
		// Begin on the front side
		sector = sidedefs[linedefs[startline].s1].sector;
		startvertex = linedefs[startline].v2;
		endvertex = linedefs[startline].v1;
		lastvertex = endvertex;
		startforward = 1;
		
		// Fill initial list
		firstlist[0] = linedefs[startline].s1;
		numlines = 1;
		longestlist = 1;
	}
	else
	{
		// Begin on the back side
		sector = sidedefs[linedefs[startline].s2].sector;
		startvertex = linedefs[startline].v1;
		endvertex = linedefs[startline].v2;
		lastvertex = endvertex;
		startforward = 0;
		
		// Fill initial list
		firstlist[0] = linedefs[startline].s2;
		numlines = 1;
		longestlist = 1;
	}
	
	// Recursively find the loop back to the startline
	int* resultlist = NextVertexToLoop(linedefs, sidedefs, startvertex, endvertex,
								startline, startside, startforward, sector, firstlist,
								numlines, lastvertex, longestlist);
	
	// Free our initial list
	delete[] firstlist;
	
	// Check if a loop was found
	if(resultlist != NULL)
	{
		// Select the lines in the loop
		for(i = 0; i < numlines; i++)
		{
			// Get linedef number
			ld = sidedefs[resultlist[i]].linedef;
			
			// Add appropriate selection code
			if(linedefs[ld].s1 == resultlist[i]) linedefs[ld].selected |= 1;
			else linedefs[ld].selected |= 2;
		}
		
		// Loop found
		return resultlist;
	}
	else
	{
		// Nothing found
		return NULL;
	}
}


// This finds all sector loops and optionally creates the array of found lines
//----------------------------------------------------------------------------
void FindSectorLoops(MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int sector, bool filllinesarray)
{
	int curlastvertex = 0;
	int curlastlength = 0;
	
	// Go for all sector linedefs
	for(int i = 0; i < numsectorlines; i++)
	{
		// Get line index and side
		int ld = sectorlines[i];
		int sd = sectorsides[i];
		
		// Check if not yet marked as used
		if((linedefs[ld].selected & sd) == 0)
		{
			// Find a loop with this line
			int resultlines = 0;
			int* result = FindLinedefLoop(linedefs, sidedefs, ld, sd - 1,
									curlastvertex, curlastlength, resultlines);
			
			// Check if open
			if(result == NULL)
			{
				// Check if longer result
				if(curlastlength > lastlength)
				{
					// Then use this lastvertex
					lastvertex = curlastvertex;
					lastlength = curlastlength;
				}
			}
			else
			{
				// Check if lines array must be created
				if(filllinesarray)
				{
					// Reallocate list array
					ReallocIntP(looplines, numlooplines, numlooplines + resultlines);
					
					// Add the results to the list
					memcpy(&(looplines[numlooplines]), result, resultlines * sizeof(int));
					numlooplines += resultlines;
				}
				
				// Destroy the result list
				delete[] result;
			}
		}
	}
}



// TestUnclosedSector: Tests if a sector is unclosed
// This returns the index to display or -1 when the sector is closed
//----------------------------------------------------------------------------
int APIENTRY TestUnclosedSector(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs,
								int numlinedefs, int sector)
{
	int outsidelinefound = 0;
	int ld;
	
	// Reset
	sectorlines = NULL;
	sectorsides = NULL;
	sectorlinescap = 0;
	numsectorlines = 0;
	lastvertex = -1;
	lastlength = 0;
	looplines = NULL;
	numlooplines = 0;
	
	
	// Find all lines on this sector
	FindSectorSideLines(linedefs, sidedefs, numlinedefs, sector);
	
	
	// Find all possible loops in this sector
	// and store them in the loop lines array
	FindSectorLoops(linedefs, sidedefs, sector, true);
	
	
	// Check if the sector is currently considered open
	if(lastvertex != -1)
	{
		// We must check if the remaining lines are not
		// completely in or on the edge of the sector boundaries
		
		// Go for all linedefs
		for(ld = 0; ld < numlinedefs; ld++)
		{
			// Check if the line has a front side
			if(linedefs[ld].s1 != -1)
			{
				// Check if not yet marked as used
				if((linedefs[ld].selected & 1) == 0)
				{
					// Check if the front side of this line goes with this sector
					if(sidedefs[linedefs[ld].s1].sector == sector)
					{
						// Test if completely within one of the boundaries.
						if(pointinsidedefs(vertices, linedefs, sidedefs, linedefs[ld].v1, looplines, numlooplines) &&
						   pointinsidedefs(vertices, linedefs, sidedefs, linedefs[ld].v2, looplines, numlooplines))
						{
							// This line lies completely in a closed boundary
							// Mark this side as done.
							linedefs[ld].selected |= 1;
							break;
						}
						else
						{
							// At least one line is not completely within boundaries,
							// then set outsidelinefound = 1 to indicate a leak
							outsidelinefound = 1;
						}
					}
				}
				
				// Check if not yet marked as used
				if((linedefs[ld].selected & 2) == 0)
				{
					// Check if the back side of this line goes with this sector
					if(sidedefs[linedefs[ld].s2].sector == sector)
					{
						// Test if completely within one of the boundaries.
						if(pointinsidedefs(vertices, linedefs, sidedefs, linedefs[ld].v1, looplines, numlooplines) &&
						   pointinsidedefs(vertices, linedefs, sidedefs, linedefs[ld].v2, looplines, numlooplines))
						{
							// This line lies completely in a closed boundary
							// Mark this side as done.
							linedefs[ld].selected |= 2;
							break;
						}
						else
						{
							// At least one line is not completely within boundaries,
							// then set outsidelinefound = 1 to indicate a leak
							outsidelinefound = 1;
						}
					}
				}
			}
		}
		
		// Second chance to find something invalid
		lastvertex = -1;
		lastlength = 0;
		
		// If an outside line is found
		if(outsidelinefound == 1)
		{
			// Find all possible loops in this sector
			FindSectorLoops(linedefs, sidedefs, sector, false);
		}
	}
	
	// Clean up the lists
	if(looplines != NULL) delete[] looplines;
	if(sectorlines != NULL) delete[] sectorlines;
	if(sectorsides != NULL) delete[] sectorsides;
	numlooplines = 0;
	sectorlinescap = 0;
	numsectorlines = 0;
	
	// Return the result
	return lastvertex;
}
