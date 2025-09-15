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
#define MAP_RENDER_SCALE			0.01
#define MAX_TRACE_SECTORS			1000

// Includes
#include <windows.h>
#include <objbase.h>
#include <wchar.h>
#include "const.h"
#include "data.h"
#include "math.h"

// turn off some stupid warnings
#pragma warning (disable: 4244)					// conversion from 'type1' to 'type2', possible loss of data


// IntersectLine: Tests a line with a ray.
// Returns -1 when not intersection, otherwise returns the distance in percent (0.0 - 1.0)
//-----------------------------------------------------------------------------
float inline IntersectLine(FPOINT3D &r1, FPOINT3D &r2, float &x3, float &y3, float &x4, float &y4)
{
	// Calculate the intersection distance from the ray
	float u_ray = ((x4 - x3) * (r1.y - y3) - (y4 - y3) * (r1.x - x3)) /
				  ((y4 - y3) * (r2.x - r1.x) - (x4 - x3) * (r2.y - r1.y));
	
	// Calculate the intersection distance from the line
	float u_line = ((r2.x - r1.x) * (r1.y - y3) - (r2.y - r1.y) * (r1.x - x3)) /
				   ((y4 - y3) * (r2.x - r1.x) - (x4 - x3) * (r2.y - r1.y));
	
	// Return the distance on the ray if intersecting, or return -1
	if((u_ray >= 0) && (u_ray <= 1) && (u_line >= 0) && (u_line <= 1)) return u_ray; else return -1;
}


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
						MAPSECTOR* sectors, MAPSSECTOR* ssectors, int* r_sidedefs, int r_numsidedefs,
						int numlinedefs, int* r_subsectors, int r_numsubsectors,
						FPOINT3D &r1, FPOINT3D &r2, int &hit_index)
{
	int ld, rs, sd, ss, s, s_found;
	int hit_type = 0;		// 0=Nothing, 1=Floor, 2=Ceiling, 3=Upper, 4=Lower, 5=Middle
	int r_sectors[MAX_TRACE_SECTORS];
	int r_numsectors = 0;
	int p_hit_type;
	float i_distance;
	float hit_distance = 10;
	float sectorheight;
	MAPSECTOR* thissector;
	MAPSECTOR* othersector;
	MAPSIDEDEF* otherside;
	FPOINT3D ip;
	
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
		if((i_distance > -1) && (i_distance < hit_distance))
		{
			// Initialize variables
			othersector = NULL;
			otherside = NULL;
			
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
				// Check if between ceiling and floor (otherwise its NOT on the wall)
				if((ip.z > thissector->hfloor) && (ip.z < thissector->hceiling))
				{
					// Check if the line is single-sided
					if(othersector == NULL)
					{
						// The point hit on the middle texture
						hit_distance = i_distance;
						hit_type = 5; // Middle
						hit_index = sd;
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
						}
						// Check if below the other sector's floor
						else if(ip.z < othersector->hfloor)
						{
							// The point hit on the lower texture
							hit_distance = i_distance;
							hit_type = 4; // Lower
							hit_index = sd;
						}
						// Check if a valid middle texture is set
						else if((sidedefs[sd].middle != NULL) && (sidedefs[sd].middle[0] != '-'))
						{
							// The point hit on the middle texture
							hit_distance = i_distance;
							hit_type = 5; // Middle
							hit_index = sd;
						}
					}
				}
			}
		}
	}
	
	// Go for all visible subsectors to test its sector
	for(ss = 0; ss < r_numsubsectors; ss++)
	{
		// Check if this subsectors sector is already tested
		s_found = false;
		for(s = 0; s < r_numsectors; s++)
		{
			// Check if this subsector is already added
			if(r_sectors[s] == ssectors[r_subsectors[ss]].sector)
			{
				s_found = true;
				break;
			}
		}
		
		// Check if already tested
		if(!s_found)
		{
			// Add sector to the list
			if(r_numsubsectors < MAX_TRACE_SECTORS)
			{
				r_sectors[r_numsectors] = ssectors[r_subsectors[ss]].sector;
				r_numsectors++;
			}
			
			// Get the sector
			MAPSECTOR* sector = &sectors[ssectors[r_subsectors[ss]].sector];
			
			// Check if the ray starts above the floor and ends below
			if((r1.z > sector->hfloor) && (r2.z < sector->hfloor))
			{
				// An intersection will be with the floor
				sectorheight = sector->hfloor;
				s_found = true;
				p_hit_type = 1; // Floor
			}
			// Check if the ray starts below the ceiling and ends above
			else if((r1.z < sector->hceiling) && (r2.z > sector->hceiling))
			{
				// An intersection will be with the ceiling
				sectorheight = sector->hceiling;
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
				// Get the distance to possible intersection
				i_distance = (sectorheight - r1.z) / (r2.z - r1.z);
				
				// Check if the distance is within the ray and closer than previous hit
				if((i_distance >= 0) && (i_distance <= 1) && (i_distance < hit_distance))
				{
					// Calculate the intersection point
					ip.x = r1.x + i_distance * (r2.x - r1.x);
					ip.y = r1.y + i_distance * (r2.y - r1.y);
					ip.z = r1.z + i_distance * (r2.z - r1.z);
					
					// Get the sector in which the intersection point lies
					int insector = IntersectSectorA(ip.x, -ip.y, vertices, linedefs, sidedefs, numlinedefs);
					
					// Check if this is the same sector
					if(insector == ssectors[r_subsectors[ss]].sector)
					{
						// Apply the hit
						hit_distance = i_distance;
						hit_type = p_hit_type;
						hit_index = insector;
					}
				}
			}
		}
	}
	
	// Return result
	return hit_type;
}


