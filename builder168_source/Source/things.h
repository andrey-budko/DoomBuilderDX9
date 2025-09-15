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


// Function declarations


// GetThingZ: Returns the Z coordinate for the given thing
//----------------------------------------------------------------------------
int inline GetThingZ(MAPSECTOR* sectors, MAPTHING* thing)
{
	int thingtop, thingbottom = 0;
	
	// This must have a sector
	if(thing->sector > 0)
	{
		// Determine thing bottom and top
		if(thing->hangs)
		{
			// Hangs from ceiling
			thingtop = sectors[thing->sector].hceiling;
			if(thing->z > 0) thingtop -= thing->z;
			thingbottom = thingtop - thing->height;
			
			// Check if below floor
			if(thingbottom < sectors[thing->sector].hfloor)
			{
				// Put against floor
				thingbottom = sectors[thing->sector].hfloor;
				thingtop = thingbottom + thing->height;
			}
		}
		else
		{
			// Stands on floor
			thingbottom = sectors[thing->sector].hfloor;
			if(thing->z > 0) thingbottom += thing->z;
			thingtop = thingbottom + thing->height;
			
			// Check if above ceiling
			if(thingtop > sectors[thing->sector].hceiling)
			{
				// Put against ceiling
				thingtop = sectors[thing->sector].hceiling;
				thingbottom = thingtop - thing->height;
			}
		}
	}
	
	// Return result
	return thingbottom;
}

