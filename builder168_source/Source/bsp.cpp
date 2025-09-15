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
#define DEG2A(x)					(((x) * 0.01745329252f))
#define NODE_SUBSECTOR				0x8000

// Includes
#include <windows.h>
#include <objbase.h>
#include "const.h"
#include "data.h"
#include "math.h"
#include "clip.h"
#include "poly.h"
#include "things.h"

// Variables
POINT* vertices = NULL;
MAPLINEDEF* linedefs = NULL;
MAPSIDEDEF* sidedefs = NULL;
MAPSEG* segs = NULL;
MAPSECTOR* sectors = NULL;
MAPSSECTOR* ssectors = NULL;
MAPNODE* nodes = NULL;
MAPTHING* things = NULL;
int numnodes = 0;
int numsectors = 0;
int numssectors = 0;
int numthings = 0;
int* renderarray = NULL;
int ssectorinsertpos = 0;
int maxssectors = 0;
int *sidedefsarray = NULL;
int sidedefinsertpos = 0;
int maxsidedefs = 0;
int* allthingsarray = NULL;
int* allthingsdistances = NULL;
int renderdistance = 0;
int currentx, currenty;


// SetStructurePointers: Sets the structure pointers to use for BSP processing
//-----------------------------------------------------------------------------
void APIENTRY SetStructurePointers(POINT* pvertices, MAPLINEDEF* plinedefs, MAPSIDEDEF* psidedefs,
							MAPSEG* psegs, MAPSECTOR* psectors, MAPSSECTOR* pssectors,
							MAPTHING* pthings, MAPNODE* pnodes, int nnumnodes,
							int nnumsectors, int nnumssectors, int nnumthings)
{
	// Just set the pointers
	vertices = pvertices;
	linedefs = plinedefs;
	sidedefs = psidedefs;
	segs = psegs;
	sectors = psectors;
	ssectors = pssectors;
	nodes = pnodes;
	things = pthings;
	numnodes = nnumnodes;
	numsectors = nnumsectors;
	numssectors = nnumssectors;
	numthings = nnumthings;
	
	// Create things array
	allthingsarray = new int[nnumthings];
	allthingsdistances = new int[nnumthings];
}


// DestroyStructurePointers: Removes the structure pointers
//-----------------------------------------------------------------------------
void APIENTRY DestroyStructurePointers()
{
	// Clean up things array
	if(allthingsarray != NULL) delete[] allthingsarray;
	if(allthingsdistances != NULL) delete[] allthingsdistances;
	
	// Reset the pointers
	allthingsarray = NULL;
	allthingsdistances = NULL;
	vertices = NULL;
	linedefs = NULL;
	sidedefs = NULL;
	segs = NULL;
	sectors = NULL;
	ssectors = NULL;
	nodes = NULL;
	things = NULL;
	numnodes = 0;
	numsectors = 0;
	numssectors = 0;
	numthings = 0;
}


// QuickSortThings: sorts the allthingsarray and allthingsdistances
// with the quicksort algorythm by distance
//-----------------------------------------------------------------------------
void QuickSortThings(int lo, int hi)
{
	// hi is the upper index of the array
	// lo is the lower index of the array
	int i = lo, j = hi, h;
	int x = allthingsdistances[(lo + hi) / 2];
	
	// Sort partition
	do
	{    
		while(allthingsdistances[i] < x) i++; 
		while(allthingsdistances[j] > x) j--;
		if(i <= j)
		{
			// Swap distances
			h = allthingsdistances[i];
			allthingsdistances[i] = allthingsdistances[j];
			allthingsdistances[j] = h;
			
			// Swap items
			h = allthingsarray[i];
			allthingsarray[i] = allthingsarray[j];
			allthingsarray[j] = h;
			
			i++;
			j--;
		}
	}
	while (i <= j);
	
	// Recursion
	if(lo < j) QuickSortThings(lo, j);
	if(i < hi) QuickSortThings(i, hi);
}


// CreateSSectorReferences: Creates the needed references for all SSECTORs
//-----------------------------------------------------------------------------
void APIENTRY CreateSSectorReferences()
{
	int s, sg, seg;
	
	// Go for all subsectors
	for(s = 0; s < numssectors; s++)
	{
		// Set to -1 when not with a sector
		// With a good nodebuilder this should never be though
		ssectors[s].sector = -1;
		
		// Get first segment
		seg = ssectors[s].startseg;
		
		// Go for all segments
		for(sg = 0; sg < ssectors[s].numsegs; sg++)
		{
			// Determine the sector to which this sub-sector belongs
			if(segs[seg + sg].side == 0)
			{
				// Check if a sidedef is assigned here
				if(linedefs[segs[seg + sg].linedef].s1 > 1)
				{
					// Sector from sidedef
					ssectors[s].sector = sidedefs[linedefs[segs[seg + sg].linedef].s1].sector;
					break;
				}
			}
			else
			{
				// Check if a sidedef is assigned here
				if(linedefs[segs[seg + sg].linedef].s2 > 1)
				{
					// Sector from sidedef
					ssectors[s].sector = sidedefs[linedefs[segs[seg + sg].linedef].s2].sector;
					break;
				}
			}
		}
		
		// Set the number of vertices to -1 to indicate these
		// have not been generated yet (a subsector may have 0 vertices)
		ssectors[s].numvertices = -1;
	}
}


// PrepareNode: Recursively walks a node tree and makes preparations to subsectors.
//-----------------------------------------------------------------------------
void __fastcall PrepareNode(int thisnode, MAPSPLIT* splitslist, int splits,
					   MAPNODE* prevnode, int nodeside)
{
	// Check if this is a subsector
	if(thisnode & NODE_SUBSECTOR)
	{
		// Get the subsector index
		int ssi = thisnode & (~NODE_SUBSECTOR);
		
		// Get the subsector
		MAPSSECTOR* ss = &ssectors[ssi];
		
		// Check if not yet triangulated
		if(ss->numvertices < 0)
		{
			// Triangulate this subsector
			RECT bounds = TriangulateSSector(vertices, segs, ss, splitslist, splits);
			
			// Apply the new boundaries to the right side of the previous node
			if(nodeside)
			{
				prevnode->lbottom = bounds.bottom;
				prevnode->ltop = bounds.top;
				prevnode->lleft = bounds.left;
				prevnode->lright = bounds.right;
			}
			else
			{
				prevnode->rbottom = bounds.bottom;
				prevnode->rtop = bounds.top;
				prevnode->rleft = bounds.left;
				prevnode->rright = bounds.right;
			}
		}
	}
	// Otherwise this is a tree node
	else
	{
		// Get the current node
		MAPNODE* nd = &nodes[thisnode];
		
		// Make a new splits list
		MAPSPLIT* newlist = (MAPSPLIT*)malloc((splits + 1) * sizeof(MAPSPLIT));
		if(splits > 0) memcpy(newlist, splitslist, splits * sizeof(MAPSPLIT));
		
		// Add the split for left of line
		newlist[splits].x = nd->x;
		newlist[splits].y = nd->y;
		newlist[splits].dx = -nd->dx;		// Note
		newlist[splits].dy = -nd->dy;		// Must be flipped for left side
		
		// Do the left side of the line
		PrepareNode(nd->left, newlist, (splits + 1), nd, 1);
		
		// Add the split for right of line
		newlist[splits].x = nd->x;
		newlist[splits].y = nd->y;
		newlist[splits].dx = nd->dx;
		newlist[splits].dy = nd->dy;
		
		// Do the left side of the line
		PrepareNode(nd->right, newlist, (splits + 1), nd, 0);
		
		// Clean up
		free(newlist);
	}
}


// PrepareAll: Prepares nodes and subsectors.
//-----------------------------------------------------------------------------
void APIENTRY PrepareAllSSectors()
{
	PrepareNode((numnodes - 1), NULL, 0, NULL, 0);
}


// WalkNode: Recursively walks a node tree or process a subsector.
// returns true when the walking should be terminated, otherwise returns false.
//-----------------------------------------------------------------------------
int __fastcall WalkNode(int thisnode, MAPSPLIT* splitslist, int splits,
				    MAPNODE* prevnode, int nodeside)
{
	int result;
	int i, sdi;
	int sdifound;
	int ssx, ssy;
	int l, r, t, b;
	int visible;
	float lt, rt, lb, rb;
	int blocksview;
	int inside;
	
	// Check if distance check can be made
	if(prevnode)
	{
		// Check what side its on
		if(nodeside)
		{
			// Get the subsector distance from position
			ssx = prevnode->lleft + ((prevnode->lright - prevnode->lleft) >> 1);
			ssy = prevnode->ltop + ((prevnode->lbottom - prevnode->ltop) >> 1);
			
			// Leave if distance is out of render limit
			if((distance(currentx, currenty, ssx, ssy) - MAX(abs(prevnode->lright - prevnode->lleft), abs(prevnode->lbottom - prevnode->ltop))) > renderdistance) return false;
		}
		else
		{
			// Get the subsector distance from position
			ssx = prevnode->rleft + ((prevnode->rright - prevnode->rleft) >> 1);
			ssy = prevnode->rtop + ((prevnode->rbottom - prevnode->rtop) >> 1);
			
			// Leave if distance is out of render limit
			if((distance(currentx, currenty, ssx, ssy) - MAX(abs(prevnode->rright - prevnode->rleft), abs(prevnode->rbottom - prevnode->rtop))) > renderdistance) return false;
		}
	}
	
	// Check if this is a subsector
	if(thisnode & NODE_SUBSECTOR)
	{
		// Check what side its on
		if(nodeside)
		{
			// Get the bounds of this subsector
			l = prevnode->lleft;
			r = prevnode->lright;
			t = prevnode->ltop;
			b = prevnode->lbottom;
		}
		else
		{
			// Get the bounds of this subsector
			l = prevnode->rleft;
			r = prevnode->rright;
			t = prevnode->rtop;
			b = prevnode->rbottom;
		}
		
		// Get the angles of bounding box endpoints
		lt = angle(currentx, currenty, l, t) + PI;
		rt = angle(currentx, currenty, r, t) + PI;
		lb = angle(currentx, currenty, l, b) + PI;
		rb = angle(currentx, currenty, r, b) + PI;
		
		// Are we inside the bounds of this subsector?
		inside = ((currentx >= l) && (currentx <= r) && (currenty >= t) && (currenty <= b));
		
		// Do the visibility test (cross)
		visible = (TestClipRange(lt, rb) || TestClipRange(rt, lb));
		
		// Check visibility result
		if(visible || inside)
		{
			// Get the subsector index
			int ssi = thisnode & (~NODE_SUBSECTOR);
			
			// Get the subsector
			MAPSSECTOR* ss = &ssectors[ssi];
			
			// Get the first seg in the subsector
			MAPSEG* sg = &segs[ss->startseg];
			
			// Go for all segs in the subsector
			for(i = 0; i < ss->numsegs; i++)
			{
				// Check if im on the front side of this seg
				if(side_of_line(vertices[sg->v1].x, vertices[sg->v1].y,  vertices[sg->v2].x, vertices[sg->v2].y, currentx, currenty) < 0)
				{
					// Get the sidedef index
					if(sg->side) sdi = linedefs[sg->linedef].s2; else sdi = linedefs[sg->linedef].s1;
					
					// Check if a sidedef exists here
					if(sdi > -1)
					{
						// Find the sidedef in the sidedefs render list
						sdifound = false;
						for(int k = 0; k < sidedefinsertpos; k++)
						{
							if(sidedefsarray[k] == sdi) 
							{
								// Already added for rendering, forget about it
								sdifound = true;
								break;
							}
						}
						
						// Check if should be rendered and render array is not full
						if(!sdifound && (sidedefinsertpos < maxsidedefs))
						{
							// Add to render array
							sidedefsarray[sidedefinsertpos] = sdi;
							sidedefinsertpos++;
						}
						
						// Check if the wall blocks my view
						if(linedefs[sg->linedef].s2 == -1)
						{
							// Blocks view
							blocksview = 1;
						}
						// Check if we can check for closed sector ceiling/floor
						// Only do this when not inside the sector
						else if(!inside && (linedefs[sg->linedef].s1 != -1))
						{
							// Get sectors
							MAPSECTOR* sc1 = sectors + sidedefs[linedefs[sg->linedef].s1].sector;
							MAPSECTOR* sc2 = sectors + sidedefs[linedefs[sg->linedef].s2].sector;
							
							// Check if the sectors block the view
							if((sc1->hfloor >= sc1->hceiling) ||
							   (sc1->hfloor >= sc2->hceiling) ||
							   (sc2->hfloor >= sc1->hceiling) ||
							   (sc2->hfloor >= sc2->hceiling))
							{
								// Blocks view
								blocksview = 1;
							}
							else
							{
								// Doesnt block
								blocksview = 0;
							}
						}
						else
						{
							// Doesnt block
							blocksview = 0;
						}
						
						// Only write the wall to clipbuffer if it blocks my view
						if(blocksview)
						{
							// Get the angles of the seg vertices towards the current position
							float a1 = angle(currentx, currenty, vertices[sg->v1].x, vertices[sg->v1].y) + PI;
							float a2 = angle(currentx, currenty, vertices[sg->v2].x, vertices[sg->v2].y) + PI;
							
							// Write on the clip buffer
							WriteClipRange(a1, a2);
						}
					}
				}
				
				// Next seg
				sg++;
			}
			
			// Check if render array is not full
			if(ssectorinsertpos < maxssectors)
			{
				// Add the subsector to the renderarray
				renderarray[ssectorinsertpos] = ssi;
				ssectorinsertpos++;
				
				// Mark sector as visible
				sectors[ss->sector].visible = 1;
				
				// Triangulate the subsector if not done yet
				if(ss->numvertices < 0)
				{
					// Triangulate this subsector
					RECT bounds = TriangulateSSector(vertices, segs, ss, splitslist, splits);
					
					// Apply the new boundaries to the right side of the previous node
					if(nodeside)
					{
						prevnode->lbottom = bounds.bottom;
						prevnode->ltop = bounds.top;
						prevnode->lleft = bounds.left;
						prevnode->lright = bounds.right;
					}
					else
					{
						prevnode->rbottom = bounds.bottom;
						prevnode->rtop = bounds.top;
						prevnode->rleft = bounds.left;
						prevnode->rright = bounds.right;
					}
				}
			}
		}
		
		// Leave
		return false;
	}
	// Otherwise this is a tree node
	else
	{
		// Get the current node
		MAPNODE* nd = &nodes[thisnode];
		
		// Make a new splits list
		MAPSPLIT* newlist = (MAPSPLIT*)malloc((splits + 1) * sizeof(MAPSPLIT));
		if(splits > 0) memcpy(newlist, splitslist, splits * sizeof(MAPSPLIT));
		
		// Check on the left (backside) of the line
		if(side_of_line(nd->x, nd->y, nd->x + nd->dx, nd->y + nd->dy, currentx, currenty) > 0)
		{
			// Add the split for left of line
			newlist[splits].x = nd->x;
			newlist[splits].y = nd->y;
			newlist[splits].dx = -nd->dx;		// Note
			newlist[splits].dy = -nd->dy;		// Must be flipped for left side
			
			// Do the left side of the line
			result = WalkNode(nd->left, newlist, (splits + 1), nd, 1);
			
			// Check if we should continue
			if(!result)
			{
				// Add the split for right of line
				newlist[splits].x = nd->x;
				newlist[splits].y = nd->y;
				newlist[splits].dx = nd->dx;
				newlist[splits].dy = nd->dy;
				
				// Do the right side of the line
				result = WalkNode(nd->right, newlist, (splits + 1), nd, 0);
			}
		}
		// Otherwise on the right (frontside) of the line
		else
		{
			// Add the split for right of line
			newlist[splits].x = nd->x;
			newlist[splits].y = nd->y;
			newlist[splits].dx = nd->dx;
			newlist[splits].dy = nd->dy;
			
			// Do the right side of the line
			result = WalkNode(nd->right, newlist, (splits + 1), nd, 0);
			
			// Check if we should continue
			if(!result)
			{
				// Add the split for left of line
				newlist[splits].x = nd->x;
				newlist[splits].y = nd->y;
				newlist[splits].dx = -nd->dx;		// Note
				newlist[splits].dy = -nd->dy;		// Must be flipped for left side
				
				// Do the left side of the line
				result = WalkNode(nd->left, newlist, (splits + 1), nd, 1);
			}
		}
		
		// Clean up
		free(newlist);
		
		// Check if we should continue
		if(!result) result = TestFullClipbuffer();
		
		// Return result
		return result;
	}
}


// ProcessBSP: Walks throught the BSP tree structure and fills the renderarray.
// The number of ssectors added to the renderarray will be returned.
//-----------------------------------------------------------------------------
void APIENTRY ProcessBSP(int* i_renderarray, int i_maxssectors, int *i_sidedefsarray, int i_maxsidedefs,
					int &i_numssectors, int &i_numsidedefs, int x, int y, int z, float angle, int fov,
					int i_renderdistance, int* thingsarray, int &i_numthings, int maxthings)
{
	int numvisiblethings;
	int i, td;
	int x2, y2;
	int thingbottom;
	
	// Initialize clipper
	InitializeClipper();
	
	// Do not clip out with an FOV 180 or higher
	if(fov < 180)
	{
		// Backward angle
		float backangle = angle + PI;
		if(backangle > PI2) backangle -= PI2;
		
		// Calculate the angles for the FOV
		float a1 = backangle - PI + DEG2A(fov);
		if(backangle < 0) backangle += PI2;
		float a2 = backangle + PI - DEG2A(fov);
		if(backangle > PI2) backangle -= PI2;
		
		// Erase where we cant see
		WriteClipRange(a1, backangle);
		WriteClipRange(backangle, a2);
	}
	
	// Reset visible parameter for all sectors
	for(i = 0; i < numsectors; i++) sectors[i].visible = 0;
	
	// Keep array pointers and limits
	renderarray = i_renderarray;
	maxssectors = i_maxssectors;
	sidedefsarray = i_sidedefsarray;
	maxsidedefs = i_maxsidedefs;
	renderdistance = i_renderdistance;
	currentx = x;
	currenty = y;
	
	// Defaults
	ssectorinsertpos = 0;
	sidedefinsertpos = 0;
	numvisiblethings = 0;
	
	// Start walking the tree from the top (last) node.
	int result = WalkNode((numnodes - 1), NULL, 0, NULL, NULL);
	
	// Calculate second coordinates for my horizontal line
	x2 = (int)(x + sin(PI2 - angle) * 100);
	y2 = (int)(y + cos(PI2 - angle) * 100);
	
	// Now go for all things 
	for(i = 0; i < numthings; i++)
	{
		// Check if thing is in a visible sector
		if(sectors[things[i].sector].visible)
		{
			// Get bottom height of thing
			thingbottom = GetThingZ(sectors, &things[i]);
			
			// Check if thing is supposed to be shown
			if((things[i].size > 0) && (things[i].height > 0))
			{
				// Check if not inside the thing
				if((x < things[i].x - things[i].size) ||
				   (x > things[i].x + things[i].size) ||
				   (y < things[i].y - things[i].size) ||
				   (y > things[i].y + things[i].size) ||
				   (z < thingbottom) ||
				   (z > thingbottom + things[i].height))
				{
					// Get the distance
					td = (int)(distance(x, y, things[i].x, things[i].y) * 100);
					
					// Check if within view distance
					if(td < (renderdistance * 100))
					{
						// Check if in front of me
						if((side_of_line(x, y, x2, y2, things[i].x, things[i].y) > 0) || (fov >= 180))
						{
							// Add thing to array
							allthingsarray[numvisiblethings] = i;
							allthingsdistances[numvisiblethings] = td;
							numvisiblethings++;
						}
					}
				}
			}
		}
	}
	
	// Sort the array of things by distance
	if(numvisiblethings > 0) QuickSortThings(0, numvisiblethings - 1);
	
	// Copy things to given array up to the max things
	if(numvisiblethings > maxthings) numvisiblethings = maxthings;
	CopyMemory(thingsarray, allthingsarray, numvisiblethings * sizeof(int));
	//for(i = 0; i < numvisiblethings; i++) thingsarray[i] = allthingsarray[i];
	
	// Return results
	i_numthings = numvisiblethings;
	i_numssectors = ssectorinsertpos;
	i_numsidedefs = sidedefinsertpos;
	
	// Clean up
	renderarray = NULL;
	sidedefsarray = NULL;
	TerminateClipper();
}


// GetMissingEntries: makes an array with entries that are present in
// one array but not in the other
//-----------------------------------------------------------------------------
void APIENTRY GetMissingEntries(int* array1, int count1, int* array2, int count2, int* resultarray, int &resultcount)
{
	int c1, c2, value, isfound;
	
	// No results yet
	resultcount = 0;
	
	// Go for all items in array 1
	for(c1 = 0; c1 < count1; c1++)
	{
		// Get the value
		value = array1[c1];
		
		// Check if can be found in array 2
		isfound = false;
		for(c2 = 0; c2 < count2; c2++)
		{
			// Check if matches
			if(array2[c2] == value)
			{
				// Found, leave
				isfound = true;
				break;
			}
		}
		
		// If it was not found
		if(!isfound)
		{
			// Add value to result array
			resultarray[resultcount] = value;
			resultcount++;
		}
	}
}


