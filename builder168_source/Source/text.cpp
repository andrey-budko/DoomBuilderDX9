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
#include <string.h>
#include "const.h"
#include "data.h"


// All characters information
CHARDESC charinfo[256];


// SetFontChar sets a char in the table
void APIENTRY SetFontChar(unsigned char* character, float width, float height,
					       float u1, float u2, float v1, float v2)
{
	unsigned char ci = character[0];

	charinfo[ci].u1 = u1;
	charinfo[ci].u2 = u2;
	charinfo[ci].v1 = v1;
	charinfo[ci].v2 = v2;
	charinfo[ci].width = width;
	charinfo[ci].height = height;
}


// GetTextSize returns the rect Right and Bottom for a give text
SRECT APIENTRY GetTextSize(unsigned char* text, float charscale)
{
	long textlen = strlen((char*)text);
	float nwidth, nheight;
	CHARDESC& chr = charinfo[0];
	SRECT newsize = {0, 0, 0, 0};
	
	// Go for all chars
	for(long c = 0; c < textlen; c++)
	{
		// Get the chardesc
		chr = charinfo[text[c]];
		
		// Scale width and height
		nwidth = chr.width * charscale;
		nheight = chr.height * charscale;
		
		// Check size and add it
		newsize.right += nwidth;
		if(newsize.bottom < nheight) newsize.bottom = nheight;
	}
	
	// Return the rect
	return newsize;
}


// CreateText makes a polygon from text string
void APIENTRY CreateText(unsigned char* text, SRECT* pos, long halign, long valign,
						  long c_lt, long c_rt, long c_lb, long c_rb,
						  float charscale, FLATVERTEX* textvertex,
						  long screenwidth, long screenheight)
{
	float beginx, beginy, nwidth, nheight;
	CHARDESC& curchar = charinfo[0];
	
	long textlen = strlen((char*)text);
	SRECT textsize = GetTextSize(text, charscale);
	
	// Align the text horizontally
	switch(halign)
	{
		case 0: beginx = pos->left; break;
		case 1: beginx = pos->right - textsize.right; break;
		case 2: beginx = pos->left + (pos->right - pos->left - textsize.right) / 2; break;
	}
	
	// Align the text vertically
	switch(valign)
	{
		case 0: beginy = pos->top; break;
		case 1: beginy = pos->bottom - textsize.bottom; break;
		case 2: beginy = pos->top + (pos->bottom - pos->top - textsize.bottom) / 2; break;
	}
	
	// Loop through all chars
	for(long c = 0; c < textlen; c++)
	{
		// Get the chardesc
		curchar = charinfo[text[c]];
		
		// Scale the char
		nwidth = curchar.width * charscale;
		nheight = curchar.height * charscale;
		
		// Create the char
		
		// Lefttop Vertex
		textvertex->diffuse = c_lt;				// Vertex Color
		textvertex->rhw = 1;					// No Scaling by DirectX
		textvertex->tu1 = curchar.u1;
		textvertex->tv1 = curchar.v1;
		textvertex->x = beginx * (float)screenwidth;	// Convert coords to Screen Coords
		textvertex->y = beginy * (float)screenheight;
		
		// Next Vertex
		textvertex++;
		
		// Leftbottom Vertex
		textvertex->diffuse = c_lb;				// Vertex Color
		textvertex->rhw = 1;					// No Scaling by DirectX
		textvertex->tu1 = curchar.u1;
		textvertex->tv1 = curchar.v2;
		textvertex->x = beginx * (float)screenwidth;	// Convert coords to Screen Coords
		textvertex->y = (beginy + nheight) * (float)screenheight;
		
		// Next Vertex
		textvertex++;
		
		// Change BeginX for next Character
		beginx += nwidth;
		
		// Righttop Vertex
		textvertex->diffuse = c_rt;				// Vertex Color
		textvertex->rhw = 1;					// No Scaling by DirectX
		textvertex->tu1 = curchar.u2;
		textvertex->tv1 = curchar.v1;
		textvertex->x = beginx * (float)screenwidth;	// Convert coords to Screen Coords
		textvertex->y = beginy * (float)screenheight;
		
		// Next Vertex
		textvertex++;
		
		// Rightbottom Vertex
		textvertex->diffuse = c_rb;				// Vertex Color
		textvertex->rhw = 1;					// No Scaling by DirectX
		textvertex->tu1 = curchar.u2;
		textvertex->tv1 = curchar.v2;
		textvertex->x = beginx * (float)screenwidth;	// Convert coords to Screen Coords
		textvertex->y = (beginy + nheight) * (float)screenheight;
		
		// Next Vertex
		textvertex++;
	}
}
