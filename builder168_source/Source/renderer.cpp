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
#include <stdio.h>
#include "const.h"
#include "data.h"
#include "math.h"
#include "renderer.h"

// turn off some stupid warnings
#pragma warning (disable: 4244)					// conversion from 'type1' to 'type2', possible loss of data

// Map screen memory
int screenwidth, screenheight;
byte* screendata = NULL;

// Map scale modifiers
float xoffset = 0;
float yoffset = 0;
float wmul = 1;
float hmul = 1;


// sng: Returns 1 for positive, 0 for 0, -1 for negative
//----------------------------------------------------------------------------
int inline sgn(int s) { if(s) return s / abs(s); else return 0; }


// Render_Init: Sets the screen data pointer and width and height
//----------------------------------------------------------------------------
void APIENTRY Render_Init(byte* scdata, int scwidth, int scheight)
{
	// Set the map screen pointer and info
	screendata = scdata;
	screenwidth = scwidth;
	screenheight = scheight;
}


// Render_Term: Destroys screen data pointer
//----------------------------------------------------------------------------
void APIENTRY Render_Term()
{
	// Reset the map screen pointer and info
	screendata = NULL;
	screenwidth = 0;
	screenheight = 0;

	// Reset scale modifiers
	xoffset = 0;
	yoffset = 0;
	wmul = 1;
	hmul = 1;
}


// Render_Scale: Sets the screen scaling mode
//----------------------------------------------------------------------------
void APIENTRY Render_Scale(float left, float top, float zoom)
{
	// Set the map screen pointer and info
	xoffset = -left;
	yoffset = -top;
	wmul = zoom;
	hmul = zoom;
}


// Render_PSet: Draws a pixel on the screen data
//----------------------------------------------------------------------------
void inline Render_PSet(const int &x, const int &y, const byte &c)
{
	if((x > -1) && (x < screenwidth) && (y > -1) && (y < screenheight))
		screendata[y * screenwidth + x] = c;
}


// Render_PSetEx: Draws a pixel on the screen data without boundary check
//----------------------------------------------------------------------------
/*
void inline Render_PSetEx(const int &x, const int &y, const byte &c)
{
	screendata[y * screenwidth + x] = c;
}
*/


// Render_PSwitch: Switches a pixel on the screen data to the specified palette part
//----------------------------------------------------------------------------
void inline Render_PSwitch(const int &x, const int &y, const byte &high)
{
	if((x > -1) && (x < screenwidth) && (y > -1) && (y < screenheight))
	{
		int p = y * screenwidth + x;
		screendata[p] = (screendata[p] & 0x3F) | high;
	}
}


// Render_PSwitchEx: Switches a pixel on the screen data to the specified palette part without check
//----------------------------------------------------------------------------
/*
void inline Render_PSwitchEx(const int &x, const int &y, const byte &high)
{
	int p = y * screenwidth + x;
	screendata[p] = (screendata[p] & 0x3F) | high;
}
*/


// Render_Clear: Clears memory with specified color
//----------------------------------------------------------------------------
void APIENTRY Render_Clear(byte c) { Render_ClearF(c); }
void __fastcall Render_ClearF(byte c)
{
	// Clear memory with give character
	memset(screendata, c, screenwidth * screenheight);
}


// Render_Line: Renders a line using Bresenham's line-drawing algorithm
//----------------------------------------------------------------------------
void APIENTRY Render_Line(float x1, float y1, float x2, float y2, byte c) { Render_LineF(x1, y1, x2, y2, c); }
void __fastcall Render_LineF(float x1, float y1, float x2, float y2, byte c)
{
	int i;
	
	// Scale coordinates
	Scale(x1, y1);
	Scale(x2, y2);
	
	// Check if the line is outside the screen for sure.
	// This is quickly done by checking in which area both points are. When this
	// is above, below, right or left of the screen, then skip drawing the line.
	if( ((x1 < 0) && (x2 < 0)) ||
		((x1 > screenwidth) && (x2 > screenwidth)) ||
		((y1 < 0) && (y2 < 0)) ||
		((y1 > screenheight) && (y2 > screenheight)) ) return;
	
	// Clip the line and only continue when visible
	//if(clipline(x1, y1, x2, y2, screenwidth - 1, screenheight - 1))
	{
		// Correct the Y axis
		y1 = screenheight - y1 - 1;
		y2 = screenheight - y2 - 1;
		
		// Distance of the line
		int dx = x2 - x1;
		int dy = y2 - y1;
		
		// Positive (absolute) distance
		int dxabs = abs(dx);
		int dyabs = abs(dy);
		
		// Half distance
		int x = dyabs >> 1;
		int y = dxabs >> 1;
		
		// Direction
		int sdx = sgn(dx);
		int sdy = sgn(dy);
		
		// Start position
		int px = x1;
		int py = y1;
		
		// Draw first pixel
		Render_PSet(px, py, c);
		//Render_PSetEx(px, py, c);
		
		// Check if the line is more horizontal than vertical
		if(dxabs >= dyabs)
		{
			for(i = 0; i < dxabs; i++)
			{
				y += dyabs;
				if(y >= dxabs)
				{
					y -= dxabs;
					py += sdy;
				}
				px += sdx;
				
				// Draw pixel
				Render_PSet(px, py, c);
				//Render_PSetEx(px, py, c);
			}
		}
		// Else the line is more vertical than horizontal
		else
		{
			for(i = 0; i < dyabs; i++)
			{
				x += dxabs;
				if(x >= dyabs)
				{
					x -= dyabs;
					px += sdx;
				}
				py += sdy;
				
				// Draw pixel
				Render_PSet(px, py, c);
				//Render_PSetEx(px, py, c);
			}
		}
	}
}


// Render_Line: Renders a line using Bresenham's line-drawing algorithm in switching mode
//----------------------------------------------------------------------------
void APIENTRY Render_LineSwitched(float x1, float y1, float x2, float y2, byte high) { Render_LineSwitchedF(x1, y1, x2, y2, high); }
void __fastcall Render_LineSwitchedF(float x1, float y1, float x2, float y2, byte high)
{
	int i;
	
	// Scale coordinates
	Scale(x1, y1);
	Scale(x2, y2);
	
	// Check if the line is outside the screen for sure.
	// This is quickly done by checking in which area both points are. When this
	// is above, below, right or left of the screen, then skip drawing the line.
	if( ((x1 < 0) && (x2 < 0)) ||
		((x1 > screenwidth) && (x2 > screenwidth)) ||
		((y1 < 0) && (y2 < 0)) ||
		((y1 > screenheight) && (y2 > screenheight)) ) return;
	
	// Clip the line and only continue when visible
	//if(clipline(x1, y1, x2, y2, screenwidth - 1, screenheight - 1))
	{
		// Correct the Y axis
		y1 = screenheight - y1 - 1;
		y2 = screenheight - y2 - 1;
		
		// Distance of the line
		int dx = x2 - x1;
		int dy = y2 - y1;
		
		// Positive (absolute) distance
		int dxabs = abs(dx);
		int dyabs = abs(dy);
		
		// Half distance
		int x = dyabs >> 1;
		int y = dxabs >> 1;
		
		// Direction
		int sdx = sgn(dx);
		int sdy = sgn(dy);
		
		// Start position
		int px = x1;
		int py = y1;
		
		// Draw first pixel
		Render_PSwitch(px, py, high);
		//Render_PSwitchEx(px, py, high);
		
		// Check if the line is more horizontal than vertical
		if(dxabs >= dyabs)
		{
			for(i = 0; i < dxabs; i++)
			{
				y += dyabs;
				if(y >= dxabs)
				{
					y -= dxabs;
					py += sdy;
				}
				px += sdx;
				
				// Draw pixel
				Render_PSwitch(px, py, high);
				//Render_PSwitchEx(px, py, high);
			}
		}
		// Else the line is more vertical than horizontal
		else
		{
			for(i = 0; i < dyabs; i++)
			{
				x += dxabs;
				if(x >= dyabs)
				{
					x -= dyabs;
					px += sdx;
				}
				py += sdy;
				
				// Draw pixel
				Render_PSwitch(px, py, high);
				//Render_PSwitchEx(px, py, high);
			}
		}
	}
}


// Render_LinedefLine: Renders a linedef line
//----------------------------------------------------------------------------
void APIENTRY Render_LinedefLine(int x1, int y1, int x2, int y2, byte c, int sl) { Render_LinedefLineF(x1, y1, x2, y2, c, sl); }
void __fastcall Render_LinedefLineF(int x1, int y1, int x2, int y2, byte c, int sl)
{
	float ix2;
	float iy2;
	
	// Render linedef line
	Render_LineF(x1, y1, x2, y2, c);
	
	// Render indicator?
	if(sl)
	{
		// Middle of indicator line
		float lx = float(x2 - x1) * 0.5;
		float ly = float(-y2 + y1) * 0.5;
		
		// Indicator line begin coordinates
		float ix1 = x1 + lx;
		float iy1 = -y1 + ly;
		
		// Normalize slope and calculate coordinates
		float len = sqrt(lx * lx + ly * ly);
		if(len)
		{
			iy2 = iy1 + (lx / len) * sl;
			ix2 = ix1 - (ly / len) * sl;
		}
		else
		{
			ix2 = ix1;
			iy2 = iy1;
		}
		
		// Render indicator line
		Render_LineF(ix1, -iy1, ix2, -iy2, c);
	}
}


// Render_LinedefLine: Renders a linedef line in switched mode
//----------------------------------------------------------------------------
void APIENTRY Render_LinedefLineSwitched(int x1, int y1, int x2, int y2, byte high, int sl) { Render_LinedefLineSwitchedF(x1, y1, x2, y2, high, sl); }
void __fastcall Render_LinedefLineSwitchedF(int x1, int y1, int x2, int y2, byte high, int sl)
{
	float ix2;
	float iy2;
	
	// Render linedef line
	Render_LineSwitchedF(x1, y1, x2, y2, high);
	
	// Render the indicator?
	if(sl)
	{
		// Middle of indicator line
		float lx = float(x2 - x1) * 0.5;
		float ly = float(-y2 + y1) * 0.5;
		
		// Indicator line begin coordinates
		float ix1 = x1 + lx;
		float iy1 = -y1 + ly;
		
		// Normalize slope and calculate coordinates
		float len = sqrt(lx * lx + ly * ly);
		if(len)
		{
			iy2 = iy1 + (lx / len) * sl;
			ix2 = ix1 - (ly / len) * sl;
		}
		else
		{
			ix2 = ix1;
			iy2 = iy1;
		}
		
		// Render indicator line
		Render_LineSwitchedF(ix1, -iy1, ix2, -iy2, high);
	}
}


// Render_DottedLine: Renders a dotted line
//----------------------------------------------------------------------------
void APIENTRY Render_DottedLine(float x1, float y1, float x2, float y2, byte c) { Render_DottedLineF(x1, y1, x2, y2, c); }
void __fastcall Render_DottedLineF(float x1, float y1, float x2, float y2, byte c)
{
	int i;
	
	// Scale coordinates
	Scale(x1, y1);
	Scale(x2, y2);
	
	// Clip the line and only continue when visible
	//if(clipline(x1, y1, x2, y2, screenwidth - 1, screenheight - 1))
	{
		// Correct the Y axis
		y1 = screenheight - y1 - 1;
		y2 = screenheight - y2 - 1;
		
		// Distance of the line
		int dx = x2 - x1;
		int dy = y2 - y1;
		
		// Positive (absolute) distance
		int dxabs = abs(dx);
		int dyabs = abs(dy);
		
		// Half distance
		int x = dyabs >> 1;
		int y = dxabs >> 1;
		
		// Direction
		int sdx = sgn(dx);
		int sdy = sgn(dy);
		
		// Start position
		int px = x1;
		int py = y1;
		
		// Draw first pixel
		if((px + py) % 2 == 0) Render_PSet(px, py, c);
		
		// Check if the line is more horizontal than vertical
		if(dxabs >= dyabs)
		{
			for(i = 0; i < dxabs; i++)
			{
				y += dyabs;
				if(y >= dxabs)
				{
					y -= dxabs;
					py += sdy;
				}
				px += sdx;
				
				// Draw pixel
				if((px + py) % 2 == 0) Render_PSet(px, py, c);
			}
		}
		// Else the line is more vertical than horizontal
		else
		{
			for(i = 0; i < dyabs; i++)
			{
				x += dxabs;
				if(x >= dyabs)
				{
					x -= dyabs;
					px += sdx;
				}
				py += sdy;
				
				// Draw pixel
				if((px + py) % 2 == 0) Render_PSet(px, py, c);
			}
		}
	}
}


// Render_Box: Renders a node box
//----------------------------------------------------------------------------
void APIENTRY Render_Box(int x1, int y1, int d, byte c, int fill, byte fc) { Render_BoxF(x1, y1, d, c, fill, fc); }
void __fastcall Render_BoxF(int x1, int y1, int d, byte c, int fill, byte fc)
{
	// If fill is non-zero, the rectangle will be
	// filled with color fc
	
	int left, top, right, bottom;
	int x, y;
	
	// Scale coordinates
	Scale(x1, y1);
	
	// Correct the Y axis
	y1 = screenheight - y1 - 1;
	
	// Determine rect
	left = x1 - d;
	top = y1 - d;
	right = x1 + d;
	bottom = y1 + d;
	
	// Check if within window
	if((right > -1) && (left < screenwidth) && (bottom > -1) && (top < screenheight))
	{
		// Draw horizontal lines
		for(y = top; y <= bottom; y++)
		{
			Render_PSet(left, y, c);
			Render_PSet(right, y, c);
		}
		
		// Draw vertical lines
		for(x = (left + 1); x <= (right - 1); x++)
		{
			Render_PSet(x, top, c);
			Render_PSet(x, bottom, c);
		}
		
		// Check if we should fill
		if(fill)
		{
			// Go through entire rectangle
			for(x = (left + 1); x <= (right - 1); x++)
			{
				for(y = (top + 1); y <= (bottom - 1); y++)
					Render_PSet(x, y, fc);
			}
		}
	}
}


// Render_BoxSwitched: Renders a box in switched mode
//----------------------------------------------------------------------------
void APIENTRY Render_BoxSwitched(int x1, int y1, int d, byte high, int fill, byte fhigh) { Render_BoxSwitchedF(x1, y1, d, high, fill, fhigh); }
void __fastcall Render_BoxSwitchedF(int x1, int y1, int d, byte high, int fill, byte fhigh)
{
	// If fill is non-zero, the rectangle will be
	// filled with color fc
	
	int left, top, right, bottom;
	int x, y;
	
	// Scale coordinates
	Scale(x1, y1);
	
	// Correct the Y axis
	y1 = screenheight - y1 - 1;
	
	// Determine rect
	left = x1 - d;
	top = y1 - d;
	right = x1 + d;
	bottom = y1 + d;
	
	// Check if within window
	if((right > -1) && (left < screenwidth) && (bottom > -1) && (top < screenheight))
	{
		// Draw horizontal lines
		for(y = top; y <= bottom; y++)
		{
			Render_PSwitch(left, y, high);
			Render_PSwitch(right, y, high);
		}
		
		// Draw vertical lines
		for(x = (left + 1); x <= (right - 1); x++)
		{
			Render_PSwitch(x, top, high);
			Render_PSwitch(x, bottom, high);
		}
		
		// Check if we should fill
		if(fill)
		{
			// Go through entire rectangle
			for(x = (left + 1); x <= (right - 1); x++)
			{
				for(y = (top + 1); y <= (bottom - 1); y++)
					Render_PSwitch(x, y, fhigh);
			}
		}
	}
}


// Render_RectSwitched: Renders a rectangle in switched mode
//----------------------------------------------------------------------------
void APIENTRY Render_RectSwitched(int x1, int y1, int x2, int y2, byte high, int bordersize) { Render_RectSwitchedF(x1, y1, x2, y2, high, bordersize); }
void __fastcall Render_RectSwitchedF(int x1, int y1, int x2, int y2, byte high, int bordersize)
{
	int x, y;
	
	// Scale coordinates
	Scale(x1, y1);
	Scale(x2, y2);
	
	// Swap values to go one direction
	if(x1 > x2) { x = x2; x2 = x1; x1 = x; }
	if(y2 > y1) { y = y2; y2 = y1; y1 = y; }
	
	// Correct the Y axis
	y1 = screenheight - y1 - 1;
	y2 = screenheight - y2 - 1;
	
	// Draw vertical lines
	for(y = y1; y <= y2; y++)
	{
		for(x = 0; x < bordersize; x++)
		{
			Render_PSwitch(x1 + x, y, high);
			Render_PSwitch(x2 - x, y, high);
		}
	}
	
	// Draw horizontal lines
	for(x = (x1 + 1); x <= (x2 - 1); x++)
	{
		for(y = 0; y < bordersize; y++)
		{
			Render_PSwitch(x, y1 + y, high);
			Render_PSwitch(x, y2 - y, high);
		}
	}
}


// Render_Bitmap: Renders a bitmap mask with a specified color
//----------------------------------------------------------------------------
void APIENTRY Render_Bitmap(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2) { Render_BitmapF(bitmap, width, height, sx, sy, sw, sh, tx, ty, c1, c2); }
void __fastcall Render_BitmapF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2)
{
	int x, y, ox, oy;
	int halfsize = sw / 2;
	byte src;
	
	// Scale coordinates
	Scale(tx, ty);
	
	// Reposition target x and y
	tx = tx - halfsize;
	ty = ty - halfsize + 1;
	
	// Make the target rect
	RECT target = {tx, ty, tx + (sw - 1), ty + (sh - 1)};
	
	// Make the screen rect
	RECT screen = {0, 0, screenwidth - 1, screenheight - 1};
	
	// This will be the rendering rect
	RECT render;
	
	// Check if intersecting and calculate intersecting rectangle
	if(IntersectRect(&render, &screen, &target))
	{
		// Correct the Y axis
		sy = height - sy - 1;
		ty = screenheight - ty - 1;
		
		// Calculate the new source offsets with the clipped rectangle
		ox = sx + (render.left - target.left);
		oy = sy - (render.top - target.top);
		
		// Go for all render pixel in x
		for(x = 0; x <= (render.right - render.left); x++)
		{
			// Go for all render pixel in y
			for(y = 0; y <= (render.bottom - render.top); y++)
			{
				// Get source pixel color
				src = bitmap[(oy - y) * width + (ox + x)];
				
				// Draw target pixel if source pixel not zero
				if(src == 255)
					Render_PSet(render.left + x, (screenheight - render.top) - y, c1);
				else if(src > 0)
					Render_PSet(render.left + x, (screenheight - render.top) - y, c2);
			}
		}
	}
}


// Render_BitmapSwitched: Renders a bitmap mask in switching mode
//----------------------------------------------------------------------------
void APIENTRY Render_BitmapSwitched(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2) { Render_BitmapSwitchedF(bitmap, width, height, sx, sy, sw, sh, tx, ty, c1, c2); }
void __fastcall Render_BitmapSwitchedF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2)
{
	int x, y, ox, oy;
	int halfsize = sw / 2;
	byte src;
	
	// Scale coordinates
	Scale(tx, ty);
	
	// Reposition target x and y
	tx = tx - halfsize;
	ty = ty - halfsize;
	
	// Make the target rect
	RECT target = {tx, ty, tx + (sw - 1), ty + (sh - 1)};
	
	// Make the screen rect
	RECT screen = {0, 0, screenwidth - 1, screenheight - 1};
	
	// This will be the rendering rect
	RECT render;
	
	// Check if intersecting and calculate intersecting rectangle
	if(IntersectRect(&render, &screen, &target))
	{
		// Correct the Y axis
		sy = height - sy - 1;
		ty = screenheight - ty - 1;
		
		// Calculate the new source offsets with the clipped rectangle
		ox = sx + (render.left - target.left);
		oy = sy - (render.top - target.top);
		
		// Go for all render pixel in x
		for(x = 0; x <= (render.right - render.left); x++)
		{
			// Go for all render pixel in y
			for(y = 0; y <= (render.bottom - render.top); y++)
			{
				// Get source pixel color
				src = bitmap[(oy - y) * width + (ox + x)];
				
				// Draw target pixel if source pixel not zero
				if(src == 255)
					Render_PSwitch(render.left + x, (screenheight - render.top) - y, c1);
				else if(src > 0)
					Render_PSwitch(render.left + x, (screenheight - render.top) - y, c2);
			}
		}
	}
}


// Render_ScaledBitmap: Renders a bitmap mask with a specified color at pre-scaled coordinates
//----------------------------------------------------------------------------
void APIENTRY Render_ScaledBitmap(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2) { Render_ScaledBitmapF(bitmap, width, height, sx, sy, sw, sh, tx, ty, c1, c2); }
void __fastcall Render_ScaledBitmapF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2)
{
	int x, y, ox, oy;
	byte src;
	
	// Make the target rect
	RECT target = {tx, ty, tx + (sw - 1), ty + (sh - 1)};
	
	// Make the screen rect
	RECT screen = {0, 0, screenwidth - 1, screenheight - 1};
	
	// This will be the rendering rect
	RECT render;
	
	// Check if intersecting and calculate intersecting rectangle
	if(IntersectRect(&render, &screen, &target))
	{
		// Correct the Y axis
		sy = height - sy - 1;
		ty = screenheight - ty - 1;
		
		// Calculate the new source offsets with the clipped rectangle
		ox = sx + (render.left - target.left);
		oy = sy - (render.top - target.top);
		
		// Go for all render pixel in x
		for(x = 0; x <= (render.right - render.left); x++)
		{
			// Go for all render pixel in y
			for(y = 0; y <= (render.bottom - render.top); y++)
			{
				// Get source pixel color
				src = bitmap[(oy - y) * width + (ox + x)];
				
				// Draw target pixel if source pixel not zero
				if(src == 255)
					Render_PSet(render.left + x, (screenheight - render.top) - y, c1);
				else if(src > 0)
					Render_PSet(render.left + x, (screenheight - render.top) - y, c2);
			}
		}
	}
}


// Render_ScaledBitmapSwitched: Renders a bitmap mask with a specified color at pre-scaled coordinates in switching mode
//----------------------------------------------------------------------------
void APIENTRY Render_ScaledBitmapSwitched(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2) { Render_ScaledBitmapSwitchedF(bitmap, width, height, sx, sy, sw, sh, tx, ty, c1, c2); }
void __fastcall Render_ScaledBitmapSwitchedF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2)
{
	int x, y, ox, oy;
	byte src;
	
	// Make the target rect
	RECT target = {tx, ty, tx + (sw - 1), ty + (sh - 1)};
	
	// Make the screen rect
	RECT screen = {0, 0, screenwidth - 1, screenheight - 1};
	
	// This will be the rendering rect
	RECT render;
	
	// Check if intersecting and calculate intersecting rectangle
	if(IntersectRect(&render, &screen, &target))
	{
		// Correct the Y axis
		sy = height - sy - 1;
		ty = screenheight - ty - 1;
		
		// Calculate the new source offsets with the clipped rectangle
		ox = sx + (render.left - target.left);
		oy = sy - (render.top - target.top);
		
		// Go for all render pixel in x
		for(x = 0; x <= (render.right - render.left); x++)
		{
			// Go for all render pixel in y
			for(y = 0; y <= (render.bottom - render.top); y++)
			{
				// Get source pixel color
				src = bitmap[(oy - y) * width + (ox + x)];
				
				// Draw target pixel if source pixel not zero
				if(src == 255)
					Render_PSwitch(render.left + x, (screenheight - render.top) - y, c1);
				else if(src > 0)
					Render_PSwitch(render.left + x, (screenheight - render.top) - y, c2);
			}
		}
	}
}
