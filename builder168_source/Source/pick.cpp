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
#define MAP_RENDER_SCALE			0.01f
#define MAX_TRACE_SECTORS			1000

// Includes
#include <windows.h>
#include <objbase.h>
#include <wchar.h>
#include "const.h"
#include "data.h"
#include "math.h"
#include "things.h"


// IntersectSectorA: Returns the intersecting sector index
//----------------------------------------------------------------------------
int __fastcall IntersectSectorA(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs)
{
	int nld;
	
	// Lets get the nearest linedef
	int lddist = 0;
	nld = NearestLinedefF(x, y, vertices, linedefs, numlinedefs, lddist);
	
	// Only continue if any linedefs found
	if(nld > -1)
	{
		// Correct the Y axis
		y = -y;
		
		// Get the vertices
		MAPVERTEX* v1 = &vertices[linedefs[nld].v1];
		MAPVERTEX* v2 = &vertices[linedefs[nld].v2];
		
		// Check the side of the line
		if(side_of_line(v1->x, v1->y, v2->x, v2->y, x, y) < 0)
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


// PickObject: Makes the world space polygon for a subsector ceiling
//-----------------------------------------------------------------------------
int APIENTRY PickObject(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs,
				    MAPSECTOR* sectors, MAPSSECTOR* ssectors, MAPTHING* things,
				    int* r_sidedefs, int r_numsidedefs, int numlinedefs, int* r_subsectors,
				    int r_numsubsectors, int* r_things, int r_numthings,
				    FPOINT3D &r1, FPOINT3D &r2, FPOINT3D &hit_point, int &hit_index)
{
	int ld, rs, sd, s, s_found, t;
	int hit_type = 0;		// 0=Nothing, 1=Floor, 2=Ceiling, 3=Upper, 4=Lower, 5=Middle, 6=Thing
	int r_sectors[MAX_TRACE_SECTORS];
	int r_numsectors = 0;
	int p_hit_type;
	int frontsector, backsector;
	int hit_count = 0;
	float tl, tr, tt, tb;
	float d[6];
	float i_distance;
	float hit_distance = 1.01f;
	float sectorheight, zheight;
	float thingbottom, thingtop;
	MAPSECTOR* thissector;
	MAPSECTOR* othersector;
	MAPSIDEDEF* otherside;
	MAPTHING* thisthing;
	FPOINT3D ip;
	
	// Add the sector in which the ray begins
	frontsector = IntersectSectorA((int)r1.x, (int)-r1.y, vertices, linedefs, sidedefs, numlinedefs);
	if(frontsector > -1) { r_sectors[r_numsectors] = frontsector; r_numsectors++; }
	
	// Go for all rendered sidedefs
	for(rs = 0; rs < r_numsidedefs; rs++)
	{
		// Get the linedef
		ld = sidedefs[r_sidedefs[rs]].linedef;
		
		// Do an intersection test on this linedef vs the ray.
		i_distance = IntersectLine(r1, r2,
						vertices[linedefs[ld].v1].x, vertices[linedefs[ld].v1].y,
						vertices[linedefs[ld].v2].x, vertices[linedefs[ld].v2].y);
		
		// Check if intersecting and closer than previous hit
		if((i_distance > -0.01f) && (i_distance < hit_distance))
		{
			// Initialize variables
			othersector = NULL;
			otherside = NULL;
			
			// Check if there is a sidedef on the right
			if(linedefs[ld].s1 > -1)
			{
				// Get front sector number
				frontsector = sidedefs[linedefs[ld].s1].sector;
				
				// Check if sector is not yet listed
				if(int_in_array(r_sectors, r_numsectors, frontsector) == false)
				{
					// Add sector to possible intersections list
					if(r_numsectors < MAX_TRACE_SECTORS)
					{
						r_sectors[r_numsectors] = frontsector;
						r_numsectors++;
					}
				}
			}
			else
			{
				// No sector on the front
				frontsector = -1;
			}
			
			// Check if there is a sidedef on the left
			if(linedefs[ld].s2 > -1)
			{
				// Get front sector number
				backsector = sidedefs[linedefs[ld].s2].sector;
				
				// Check if sector is not yet listed
				if(int_in_array(r_sectors, r_numsectors, backsector) == false)
				{
					// Add sector to possible intersections list
					if(r_numsectors < MAX_TRACE_SECTORS)
					{
						r_sectors[r_numsectors] = backsector;
						r_numsectors++;
					}
				}
			}
			else
			{
				// No sector on the back
				backsector = -1;
			}
			
			// Calculate the intersection point
			ip.x = r1.x + i_distance * (r2.x - r1.x);
			ip.y = r1.y + i_distance * (r2.y - r1.y);
			ip.z = r1.z + i_distance * (r2.z - r1.z);
			
			// Check the side of line we are at
			if(side_of_line(vertices[linedefs[ld].v1].x, vertices[linedefs[ld].v1].y,
							vertices[linedefs[ld].v2].x, vertices[linedefs[ld].v2].y,
							r1.x, r1.y) < 0)
			{
				// Check if there is a sidedef on the right (sidedef 1)
				if(linedefs[ld].s1 > -1)
				{
					// Get the sector that this side belongs to
					thissector = &sectors[sidedefs[linedefs[ld].s1].sector];
					sd = linedefs[ld].s1;
					
					// Get that sector and side too
					if(linedefs[ld].s2 > -1) othersector = &sectors[sidedefs[linedefs[ld].s2].sector];
				}
				else
				{
					// No sidedef here
					sd = -1;
				}
			}
			else
			{
				// Check if there is a sidedef on the left (sidedef 2)
				if(linedefs[ld].s2 > -1)
				{
					// Get the sector that this side belongs to
					thissector = &sectors[sidedefs[linedefs[ld].s2].sector];
					sd = linedefs[ld].s2;
					
					// Get that sector and side too
					if(linedefs[ld].s1 > -1) othersector = &sectors[sidedefs[linedefs[ld].s1].sector];
				}
				else
				{
					// No sidedef here
					sd = -1;
				}
			}
			
			// Check if there is a sidedef on this side
			if(sd > -1)
			{
				// Check if between ceiling and floor, or, if crossing ceiling or floor
				if(((ip.z > thissector->hfloor) || (r1.z > thissector->hfloor)) &&
				   ((ip.z < thissector->hceiling) || (r1.z < thissector->hceiling)))
				{
					// Check if the line is single-sided
					if(othersector == NULL)
					{
						// The point hit on the middle texture
						hit_distance = i_distance;
						hit_type = 5; // Middle
						hit_index = sd;
						hit_point = ip;
						if(hit_count == 0) hit_count++; else break;
					}
					else
					{
						// Check if above the other sector's ceiling
						if(ip.z > othersector->hceiling)
						{
							// The point hit on the upper texture
							hit_distance = i_distance;
							hit_type = 3; // Upper
							hit_index = sd;
							hit_point = ip;
							if(hit_count == 0) hit_count++; else break;
						}
						// Check if below the other sector's floor
						else if(ip.z < othersector->hfloor)
						{
							// The point hit on the lower texture
							hit_distance = i_distance;
							hit_type = 4; // Lower
							hit_index = sd;
							hit_point = ip;
							if(hit_count == 0) hit_count++; else break;
						}
						// Check if a valid middle texture is set
						else if((sidedefs[sd].middle != NULL) && (sidedefs[sd].middle[0] != '-'))
						{
							// Check if point lies exactly on the texture
							if((ip.z <= sidedefs[sd].middletop) && (ip.z >= sidedefs[sd].middlebottom))
							{
								// The point hit on the middle texture
								hit_distance = i_distance;
								hit_type = 5; // Middle
								hit_index = sd;
								hit_point = ip;
								if(hit_count == 0) hit_count++; else break;
							}
						}
					}
				}
			}
		}
	}
	
	// Go for all rendered sectors to test
	for(s = 0; s < r_numsectors; s++)
	{
		// Get the sector
		MAPSECTOR* sector = &sectors[r_sectors[s]];
		
		// Check if the ray starts above the floor and ends below
		if((r1.z > sector->hfloor) && (r2.z < sector->hfloor))
		{
			// An intersection will be with the floor
			sectorheight = (float)sector->hfloor;
			s_found = true;
			p_hit_type = 1; // Floor
		}
		// Check if the ray starts below the ceiling and ends above
		else if((r1.z < sector->hceiling) && (r2.z > sector->hceiling))
		{
			// An intersection will be with the ceiling
			sectorheight = (float)sector->hceiling;
			s_found = true;
			p_hit_type = 2; // Ceiling
		}
		else
		{
			// No intersection possible
			s_found = false;
		}
		
		// Check if an intersection is possible
		if(s_found)
		{
			// Check if not aiming straight forward (then we can never hit the floor nor the ceiling)
			if(abs(r2.z - r1.z) > 0.1f)
			{
				// Get the distance to possible intersection
				i_distance = (sectorheight - r1.z) / (r2.z - r1.z);
				
				// Check if the distance is within the ray and closer than previous hit
				if((i_distance >= 0.0f) && (i_distance <= 1.0f) && (i_distance < hit_distance))
				{
					// Calculate the intersection point
					ip.x = r1.x + i_distance * (r2.x - r1.x);
					ip.y = r1.y + i_distance * (r2.y - r1.y);
					ip.z = r1.z + i_distance * (r2.z - r1.z);
					
					// Get the sector in which the intersection point lies
					int insector = IntersectSectorA((int)ip.x, (int)-ip.y, vertices, linedefs, sidedefs, numlinedefs);
					
					// Check if this is the same sector
					if(insector == r_sectors[s])
					{
						// Apply the hit
						hit_distance = i_distance;
						hit_type = p_hit_type;
						hit_index = insector;
						hit_point = ip;
					}
				}
			}
		}
	}
	
	// Now go for all rendered things to test
	for(t = 0; t < r_numthings; t++)
	{
		// Get the thing
		thisthing = &things[r_things[t]];
		
		// Check if the thing is in a sector
		// You never know :)
		if(thisthing->sector < 0) continue;
		
		// Setup square coordinates
		tl = float(thisthing->x - thisthing->size);
		tr = float(thisthing->x + thisthing->size);
		tt = float(thisthing->y - thisthing->size);
		tb = float(thisthing->y + thisthing->size);
		
		// Determine top and bottom
		thingbottom = (float)GetThingZ(sectors, thisthing);
		thingtop = thingbottom + (float)thisthing->height;
		
		// Test if the ray is a potential collision
		if(!(((r1.x < tl) && (r2.x < tl)) ||
			((r1.x > tr) && (r2.x > tr)) ||
			((r1.y < tt) && (r2.y < tt)) ||
			((r1.y > tb) && (r2.y > tb)) ||
			((r1.z < thingbottom) && (r2.z < thingbottom)) ||
			((r1.z > thingtop) && (r2.z > thingtop))) )
		{
			// Perform line vs line intersection tests
			// First 4 tests are on the XY plane while the last 2 are on the XZ plane
			d[0] = IntersectLine(r1, r2, tl, tt, tr, tt);
			d[1] = IntersectLine(r1, r2, tr, tt, tr, tb);
			d[2] = IntersectLine(r1, r2, tr, tb, tl, tb);
			d[3] = IntersectLine(r1, r2, tl, tb, tl, tt);
			d[4] = IntersectLineXZ(r1, r2, tl, thingtop, tr, thingtop);
			d[5] = IntersectLineXZ(r1, r2, tl, thingbottom, tr, thingbottom);
			
			// Go for all intersection results done on the XY plane
			for(s = 0; s < 4; s++)
			{
				// Check if the distance is smaller
				// than previous found distance
				if((d[s] > -0.1f) && (d[s] < hit_distance))
				{
					// Calculate the Z height at
					// the point of interesection
					zheight = r1.z + d[s] * (r2.z - r1.z);
					
					// Check if the intersection Z height
					// is within the thing boundaries
					if((zheight >= thingbottom) && (zheight <= thingtop))
					{
						// Yay, thing intersection!
						hit_distance = d[s];
						hit_type = 6;	// Thing
						hit_index = r_things[t];
					}
				}
			}
			
			// Go for all intersection results done on the XZ plane
			for(s = 4; s < 6; s++)
			{
				// Check if the distance is smaller
				// than previous found distance
				if((d[s] > -0.1f) && (d[s] < hit_distance))
				{
					// Calculate the Y height at
					// the point of interesection
					zheight = r1.y + d[s] * (r2.y - r1.y);
					
					// Check if the intersection Y height
					// is within the thing boundaries
					if((zheight <= tb) && (zheight >= tt))
					{
						// Yay, thing intersection!
						hit_distance = d[s];
						hit_type = 6;	// Thing
						hit_index = r_things[t];
					}
				}
			}
		}
	}
	
	// Return result
	return hit_type;
}

