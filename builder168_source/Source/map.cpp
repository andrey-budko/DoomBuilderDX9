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
#include <string>
#include "data.h"
#include "math.h"
#include "const.h"
#include "renderer.h"

// turn off some stupid warnings
#pragma warning (disable: 4244)					// conversion from 'type1' to 'type2', possible loss of data


// Render_AllLinedefs: Renders all linedefs according to their properties
//----------------------------------------------------------------------------
void APIENTRY Render_AllLinedefs(MAPVERTEX* vertices, MAPLINEDEF* linedefs, int startindex, int endindex, int submode, int indicatorlength)
{
	int sc, ld; //, l;
	MAPLINEDEF* linedef = NULL;
	
	// Determine the loop to perform
	if((submode == ESM_DRAGGING) || (submode == ESM_DRAWING))
	{
		// Go for all linedefs
		for(ld = startindex; ld <= endindex; ld++)
		{
			// Get a pointer to the linedef
			linedef = linedefs + ld;
			
			// Check if selected
			if(linedef->selected)
			{
				// Selection color
				sc = CLR_LINESELECTED;
			}
			else
			{
				// Determine impassable
				if((linedef->flags & LDF_IMPASSIBLE) || ((linedef->flags & LDF_TWOSIDED) == 0))
				{
					// Determine linedef color
					if(linedef->effect != 0)
					{
						// Impassable line with effect
						sc = CLR_LINESPECIAL;
					}
					else
					{
						// Impassable line
						sc = CLR_LINE;
					}
				}
				else
				{
					// Determine linedef color
					if(linedef->effect != 0)
					{
						// Double line with effect
						sc = CLR_LINESPECIALDOUBLE;
					}
					else if(linedef->flags & LDF_BLOCKSOUND)
					{
						// Double line blocks sound
						sc = CLR_LINEBLOCKSOUND;
					}
					else
					{
						// Double line
						sc = CLR_LINEDOUBLE;
					}
				}
			}
			
			// Render the linedef
			Render_LinedefLineF(vertices[linedef->v1].x, vertices[linedef->v1].y, vertices[linedef->v2].x, vertices[linedef->v2].y, sc, indicatorlength);
		}
	}
	else
	{
		// Go for all linedefs
		for(ld = startindex; ld <= endindex; ld++)
		{
			// Get a pointer to the linedef
			linedef = linedefs + ld;
			
			// Check if selected
			if(linedef->selected)
			{
				// Selection color
				sc = CLR_LINESELECTED;
			}
			else
			{
				// Determine impassable
				if((linedef->flags & LDF_IMPASSIBLE) || ((linedef->flags & LDF_TWOSIDED) == 0))
				{
					// Determine linedef color
					if(linedef->effect != 0)
					{
						// Impassable line with effect
						sc = CLR_LINESPECIAL;
					}
					else
					{
						// Impassable line
						sc = CLR_LINE;
					}
				}
				else
				{
					// Determine linedef color
					if(linedef->effect != 0)
					{
						// Double line with effect
						sc = CLR_LINESPECIALDOUBLE;
					}
					else if(linedef->flags & LDF_BLOCKSOUND)
					{
						// Double line blocks sound
						sc = CLR_LINEBLOCKSOUND;
					}
					else
					{
						// Double line
						sc = CLR_LINEDOUBLE;
					}
				}
			}
			
			// Render the linedef
			Render_LinedefLineF(vertices[linedef->v1].x, vertices[linedef->v1].y, vertices[linedef->v2].x, vertices[linedef->v2].y, sc, indicatorlength);
		}
	}
}


// Render_TaggedLinedefs: Renders all linedefs with the given tag
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedLinedefs(MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int argtag, int argmark, byte c, int indicatorlength, int rendervertices, int vertexsize)
{
	int ld, sc;
	MAPLINEDEF* ldp;
	
	// Go for all linedefs
	for(ld = 0; ld < numlinedefs; ld++)
	{
		// Does this linedef have an action at all?
		if(linedefs[ld].effect)
		{
			// Get a pointer to the linedef
			ldp = &linedefs[ld];
			
			// Check if any of the arguments match the tag
			if((ldp->tag == argtag) ||
			   ((ldp->arg[0] == argtag) && (ldp->argref[0] == argmark)) ||
			   ((ldp->arg[1] == argtag) && (ldp->argref[1] == argmark)) ||
			   ((ldp->arg[2] == argtag) && (ldp->argref[2] == argmark)) ||
			   ((ldp->arg[3] == argtag) && (ldp->argref[3] == argmark)) ||
			   ((ldp->arg[4] == argtag) && (ldp->argref[4] == argmark)))
			{
				// Check if we should use given color
				if(c)
				{
					sc = c;
				}
				else
				{
					// Check if selected
					if(ldp->selected)
					{
						// Selection color
						sc = CLR_LINESELECTED;
					}
					else
					{
						// Determine impassable
						if((ldp->flags & LDF_IMPASSIBLE) || ((ldp->flags & LDF_TWOSIDED) == 0))
						{
							// Determine linedef color
							if(ldp->effect != 0)
							{
								// Impassable line with effect
								sc = CLR_LINESPECIAL;
							}
							else
							{
								// Impassable line
								sc = CLR_LINE;
							}
						}
						else
						{
							// Determine linedef color
							if(ldp->effect != 0)
							{
								// Double line with effect
								sc = CLR_LINESPECIALDOUBLE;
							}
							else if(ldp->flags & LDF_BLOCKSOUND)
							{
								// Double line blocks sound
								sc = CLR_LINEBLOCKSOUND;
							}
							else
							{
								// Double line
								sc = CLR_LINEDOUBLE;
							}
						}
					}
				}
				
				// Render linedef and vertices
				Render_LinedefLineF(vertices[ldp->v1].x, vertices[ldp->v1].y, vertices[ldp->v2].x, vertices[ldp->v2].y, sc, indicatorlength);
				if(rendervertices)
				{
					Render_BoxF(vertices[ldp->v1].x, vertices[ldp->v1].y, vertexsize, CLR_VERTEX, 1, CLR_VERTEX);
					Render_BoxF(vertices[ldp->v2].x, vertices[ldp->v2].y, vertexsize, CLR_VERTEX, 1, CLR_VERTEX);
				}
			}
		}
	}
}


// Render_AllImpassableLinedefs: Renders all linedefs marked impassable
//----------------------------------------------------------------------------
void APIENTRY Render_AllImpassableLinedefs(MAPVERTEX* vertices, MAPLINEDEF* linedefs, int startindex, int endindex, int indicatorlength)
{
	int sc, ld; //, l;
	MAPLINEDEF* linedef = NULL;
	
	// Go for all linedefs
	for(ld = startindex; ld <= endindex; ld++)
	{
		// Get a pointer to the linedef
		linedef = linedefs + ld;
		
		// Check if selected
		if(linedef->selected)
		{
			// Selection color
			sc = CLR_LINESELECTED;
		}
		else
		{
			// Determine impassable
			if((linedef->flags & LDF_IMPASSIBLE) || ((linedef->flags & LDF_TWOSIDED) == 0))
			{
				// Determine linedef color
				if(linedef->effect != 0) sc = CLR_LINESPECIAL; else sc = CLR_LINE;
			}
			else
			{
				// Dont render
				sc = CLR_BACKGROUND;
			}
		}
		
		// Render the linedef
		if(sc != CLR_BACKGROUND) Render_LinedefLineF(vertices[linedef->v1].x, vertices[linedef->v1].y, vertices[linedef->v2].x, vertices[linedef->v2].y, sc, indicatorlength);
	}
}


// Render_TaggedSectors: Renders all sectors with the given tag
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedSectors(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, MAPSECTOR* sectors, int numsectors, int numlinedefs, int sectortag, byte c, int indicatorlength, int rendervertices, int vertexsize)
{
	int s, sc, ld, found;
	MAPLINEDEF* ldp = NULL;
	
	// Go for all sectors
	for(s = 0; s < numsectors; s++)
	{
		// Check if this sector tag matches
		if(sectors[s].tag == sectortag)
		{
			// Go for all linedefs
			for(ld = 0; ld < numlinedefs; ld++)
			{
				// Get a pointer to the linedef
				ldp = &linedefs[ld];
				
				// Check if any of the sidedefs belong to this sector
				found = 0;
				if(ldp->s1 > -1) if(sidedefs[ldp->s1].sector == s) found = 1;
				if(ldp->s2 > -1) if(sidedefs[ldp->s2].sector == s) found = 1;
				
				// Render if we should render this linedef
				if(found)
				{
					// Check if we should use given color
					if(c)
					{
						sc = c;
					}
					else
					{
						// Check if selected
						if(ldp->selected)
						{
							// Selection color
							sc = CLR_LINESELECTED;
						}
						else
						{
							// Determine impassable
							if((ldp->flags & LDF_IMPASSIBLE) || ((ldp->flags & LDF_TWOSIDED) == 0))
							{
								// Determine linedef color
								if(ldp->effect != 0)
								{
									// Impassable line with effect
									sc = CLR_LINESPECIAL;
								}
								else
								{
									// Impassable line
									sc = CLR_LINE;
								}
							}
							else
							{
								// Determine linedef color
								if(ldp->effect != 0)
								{
									// Double line with effect
									sc = CLR_LINESPECIALDOUBLE;
								}
								else if(ldp->flags & LDF_BLOCKSOUND)
								{
									// Double line blocks sound
									sc = CLR_LINEBLOCKSOUND;
								}
								else
								{
									// Double line
									sc = CLR_LINEDOUBLE;
								}
							}
						}
					}
					
					// Render linedef and vertices
					Render_LinedefLineF(vertices[ldp->v1].x, vertices[ldp->v1].y, vertices[ldp->v2].x, vertices[ldp->v2].y, sc, indicatorlength);
					if(rendervertices)
					{
						Render_BoxF(vertices[ldp->v1].x, vertices[ldp->v1].y, vertexsize, CLR_VERTEX, 1, CLR_VERTEX);
						Render_BoxF(vertices[ldp->v2].x, vertices[ldp->v2].y, vertexsize, CLR_VERTEX, 1, CLR_VERTEX);
					}
				}
			}
		}
	}
}


// Render_TaggedThings: Renders all things with the given tag
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedThings(MAPTHING* things, int numthings, int thingtag, byte* thingbitmaps, int bitmapswidth, int imagesize, int outlines, float outlinezoom, int filterthings, THINGFILTERS* filter)
{
	// Check if outlines must be rendered
	if(outlines)
	{
		// Go for all things
		for(int th = 0; th < numthings; th++)
		{
			// Check if this thing has this tag
			if(things[th].tag == thingtag)
			{
				// Render the outline with tag highlight color
				Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), CLR_THINGTAG, 0, 0);
			}
		}
	}
	
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Check if this thing is selected
		if(things[th].tag == thingtag)
		{
			// Render the thing with tag highlight color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, CLR_THINGTAG, CLR_BACKGROUND);
		}
	}
}


// Render_TaggedThingsNormal: Renders all things with the given tag in own color
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedThingsNormal(MAPTHING* things, int numthings, int thingtag, byte* thingbitmaps, int bitmapswidth, int imagesize, int outlines, float outlinezoom, int filterthings, THINGFILTERS* filter)
{
	// Check if outlines must be rendered
	if(outlines)
	{
		// Go for all things
		for(int th = 0; th < numthings; th++)
		{
			// Check if this thing has this tag
			if(things[th].tag == thingtag)
			{
				// Render the outline with tag highlight color
				Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), things[th].color, 0, 0);
			}
		}
	}
	
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Check if this thing is selected
		if(things[th].tag == thingtag)
		{
			// Render the thing with tag highlight color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, things[th].color, CLR_BACKGROUND);
		}
	}
}


// Render_TaggedArgThings: Renders all things with the given tag
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedArgThings(MAPTHING* things, int numthings, int argtag, int argmark, byte* thingbitmaps, int bitmapswidth, int imagesize, int outlines, float outlinezoom, int filterthings, THINGFILTERS* filter)
{
	// Check if outlines must be rendered
	if(outlines)
	{
		// Go for all things
		for(int th = 0; th < numthings; th++)
		{
			// Check if this thing has this tag
			if(((things[th].arg[0] == argtag) && (things[th].argref[0] == argmark)) ||
			   ((things[th].arg[1] == argtag) && (things[th].argref[1] == argmark)) ||
			   ((things[th].arg[2] == argtag) && (things[th].argref[2] == argmark)) ||
			   ((things[th].arg[3] == argtag) && (things[th].argref[3] == argmark)) ||
			   ((things[th].arg[4] == argtag) && (things[th].argref[4] == argmark)))
			{
				// Render the outline with tag highlight color
				Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), CLR_THINGTAG, 0, 0);
			}
		}
	}
	
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Check if this thing has this tag
		if(((things[th].arg[0] == argtag) && (things[th].argref[0] == argmark)) ||
		   ((things[th].arg[1] == argtag) && (things[th].argref[1] == argmark)) ||
		   ((things[th].arg[2] == argtag) && (things[th].argref[2] == argmark)) ||
		   ((things[th].arg[3] == argtag) && (things[th].argref[3] == argmark)) ||
		   ((things[th].arg[4] == argtag) && (things[th].argref[4] == argmark)))
		{
			// Render the thing with tag highlight color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, CLR_THINGTAG, CLR_BACKGROUND);
		}
	}
}


// Render_TaggedArgThingsNormal: Renders all things with the given tag in own color
//----------------------------------------------------------------------------
void APIENTRY Render_TaggedArgThingsNormal(MAPTHING* things, int numthings, int argtag, int argmark, byte* thingbitmaps, int bitmapswidth, int imagesize, int outlines, float outlinezoom, int filterthings, THINGFILTERS* filter)
{
	// Check if outlines must be rendered
	if(outlines)
	{
		// Go for all things
		for(int th = 0; th < numthings; th++)
		{
			// Check if this thing has this tag
			if(((things[th].arg[0] == argtag) && (things[th].argref[0] == argmark)) ||
			   ((things[th].arg[1] == argtag) && (things[th].argref[1] == argmark)) ||
			   ((things[th].arg[2] == argtag) && (things[th].argref[2] == argmark)) ||
			   ((things[th].arg[3] == argtag) && (things[th].argref[3] == argmark)) ||
			   ((things[th].arg[4] == argtag) && (things[th].argref[4] == argmark)))
			{
				// Render the outline with tag highlight color
				Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), things[th].color, 0, 0);
			}
		}
	}
	
	// Go for all things
	for(int th = 0; th < numthings; th++)
	{
		// Check if this thing has this tag
		if(((things[th].arg[0] == argtag) && (things[th].argref[0] == argmark)) ||
		   ((things[th].arg[1] == argtag) && (things[th].argref[1] == argmark)) ||
		   ((things[th].arg[2] == argtag) && (things[th].argref[2] == argmark)) ||
		   ((things[th].arg[3] == argtag) && (things[th].argref[3] == argmark)) ||
		   ((things[th].arg[4] == argtag) && (things[th].argref[4] == argmark)))
		{
			// Render the thing with tag highlight color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, things[th].color, CLR_BACKGROUND);
		}
	}
}


// Render_AllVertices: Renders all vertices according to their properties
//----------------------------------------------------------------------------
void APIENTRY Render_AllVertices(MAPVERTEX* vertices, int startindex, int endindex, int vertexsize)
{
	// Go for all vertices
	for(int v = startindex; v <= endindex; v++)
	{
		// Check if this vertex is selected
		if(vertices[v].selected)
		{
			// Draw vertex with selection color
			Render_BoxF(vertices[v].x, vertices[v].y, vertexsize, CLR_VERTEXSELECTED, 1, CLR_VERTEXSELECTED);
		}
		else
		{
			// Draw vertex with normal color
			Render_BoxF(vertices[v].x, vertices[v].y, vertexsize, CLR_VERTEX, 1, CLR_VERTEX);
		}
	}
}


// Render_AllThings: Renders all things according to their properties
//----------------------------------------------------------------------------
void APIENTRY Render_AllThings(MAPTHING* things, int startindex, int endindex, byte* thingbitmaps, int bitmapswidth, int imagesize, int outlines, float outlinezoom, int filterthings, THINGFILTERS* filter)
{
	// Check if outlines must be rendered
	if(outlines)
	{
		// Go for all things
		for(int th = startindex; th <= endindex; th++)
		{
			// Check if this thing is selected
			if(things[th].selected)
			{
				// Render the outline with selection color
				Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), CLR_THINGSELECTED, 0, 0);
			}
			else
			{
				// Check if any of the thing flags match any of the filter flags
				if(ThingFiltered(&things[th], filterthings, filter))
				{
					// Render the outline with normal color
					Render_BoxF(things[th].x, things[th].y, floor(things[th].size * outlinezoom + 0.5), things[th].color, 0, 0);
				}
			}
		}
	}
	
	// Go for all things
	for(int th = startindex; th <= endindex; th++)
	{
		// Check if this thing is selected
		if(things[th].selected)
		{
			// Render the thing with selection color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, CLR_THINGSELECTED, CLR_BACKGROUND);
		}
		else
		{
			// Check if any of the thing flags match any of the filter flags
			if(ThingFiltered(&things[th], filterthings, filter))
			{
				// Render the thing with normal color
				Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, things[th].color, CLR_BACKGROUND);
			}
		}
	}
}


// Render_AllThingsDarkened: Renders all things according to their properties, but darker
//----------------------------------------------------------------------------
void APIENTRY Render_AllThingsDarkened(MAPTHING* things, int startindex, int endindex, byte* thingbitmaps, int bitmapswidth, int imagesize, int filterthings, THINGFILTERS* filter)
{
	// Go for all things
	for(int th = startindex; th <= endindex; th++)
	{
		// Check if any of the thing flags match any of the filter flags
		if(ThingFiltered(&things[th], filterthings, filter))
		{
			// Render the thing with darker color
			Render_BitmapF(thingbitmaps, bitmapswidth, imagesize, things[th].image * imagesize, 0, imagesize, imagesize, things[th].x, things[th].y, things[th].color + 16, CLR_BACKGROUND);
		}
	}
}


// Render_ChangingLengths: Renders all numbers for changing lines
//----------------------------------------------------------------------------
void APIENTRY Render_ChangingLengths(MAPVERTEX* vertices, MAPLINEDEF* linedefs, int* changelines, int numchangelines, byte* bitmap, int width, int height, int charwidth, int charheight)
{
	int c, cx, cy;
	int lx, ly;
	int x1, y1, x2, y2;
	int l, ld;
	float len;
	char str_number[16];
	
	// Go for all changing lines
	for(l = 0; l < numchangelines; l++)
	{
		// Get lindef index
		ld = changelines[l];
		
		// Get the line coordinates
		x1 = vertices[linedefs[ld].v1].x;
		y1 = vertices[linedefs[ld].v1].y;
		x2 = vertices[linedefs[ld].v2].x;
		y2 = vertices[linedefs[ld].v2].y;
		
		// Line distances
		lx = x2 - x1;
		ly = y2 - y1;
		
		// Length of line
		len = sqrt((float)(lx * lx + ly * ly));
		
		// Round the length
		len = floor(len + 0.5);
		
		// Make a string from the length number
		sprintf(str_number, "%i", (int)len);
		
		// Scale coordinates
		Scale(x1, y1);
		Scale(x2, y2);
		
		// Half line distances
		lx = (x2 - x1) >> 1;
		ly = (y2 - y1) >> 1;
		
		// Start offset for numbers
		cx = x1 + (lx - ((charwidth * strlen(str_number)) >> 1));
		cy = y1 + (ly - (charheight >> 1));
		
		// Go for each character
		for(c = 0; str_number[c] != 0; c++)
		{
			// Draw character bitmap
			Render_ScaledBitmapF(bitmap, width, height, (str_number[c] - 48) * charwidth, 0,
						         charwidth, charheight, cx, cy, CLR_LINEHIGHLIGHT, CLR_BACKGROUND);
			
			// Change the offset for next character
			cx += (charwidth - 1);
		}
	}
}


// Render_NumberSwitched: Renders a given number with switching palette
//----------------------------------------------------------------------------
void APIENTRY Render_NumberSwitched(int number, int x, int y, byte* bitmap, int width, int height, int charwidth, int charheight, byte c1, byte c2)
{
	int c, cx, cy;
	char str_number[16];
	
	// Make a string from the number
	sprintf(str_number, "%i", number);
	
	// Scale coordinates
	Scale(x, y);
	
	// Start offset for numbers
	cx = x - ((charwidth * strlen(str_number)) >> 1);
	cy = y - (charheight >> 1);
	
	// Go for each character
	for(c = 0; str_number[c] != 0; c++)
	{
		// Draw character bitmap
		Render_ScaledBitmapSwitchedF(bitmap, width, height, (str_number[c] - 48) * charwidth, 0,
						             charwidth, charheight, cx, cy, c1, c2);
		
		// Change the offset for next character
		cx += (charwidth - 1);
	}
}


