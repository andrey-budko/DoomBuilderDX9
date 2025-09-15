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


// Map scale modifiers
extern float xoffset;
extern float yoffset;
extern float wmul;
extern float hmul;


// Scale: Scales an x and y
//----------------------------------------------------------------------------
inline void Scale(int &x, int &y)
{
	// Scale the x and y
	x = (int)(((float)x + xoffset) * wmul);
	y = (int)((-(float)y + yoffset) * hmul);
}


// Scale: Scales an x and y
//----------------------------------------------------------------------------
inline void Scale(float &x, float &y)
{
	// Scale the x and y
	x = (x + xoffset) * wmul;
	y = (-y + yoffset) * hmul;
}


void __fastcall Render_ClearF(byte c);
void __fastcall Render_LineF(float x1, float y1, float x2, float y2, byte c);
void __fastcall Render_LineSwitchedF(float x1, float y1, float x2, float y2, byte high);
void __fastcall Render_LinedefLineF(int x1, int y1, int x2, int y2, byte c, int sl);
void __fastcall Render_LinedefLineSwitchedF(int x1, int y1, int x2, int y2, byte high, int sl);
void __fastcall Render_DottedLineF(float x1, float y1, float x2, float y2, byte c);
void __fastcall Render_BoxF(int x1, int y1, int d, byte c, int fill, byte fc);
void __fastcall Render_BoxSwitchedF(int x1, int y1, int d, byte high, int fill, byte fhigh);
void __fastcall Render_RectSwitchedF(int x1, int y1, int x2, int y2, byte high, int bordersize);
void __fastcall Render_BitmapF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2);
void __fastcall Render_BitmapSwitchedF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2);
void __fastcall Render_ScaledBitmapF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2);
void __fastcall Render_ScaledBitmapSwitchedF(byte* bitmap, int width, int height, int sx, int sy, int sw, int sh, int tx, int ty, byte c1, byte c2);

