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


// Palette organization
enum ENUM_PALETTECOLORS
{
	CLR_BACKGROUND,
	CLR_VERTEX,
	CLR_VERTEXSELECTED,
	CLR_VERTEXHIGHLIGHT,
	CLR_LINE,
	CLR_LINEDOUBLE,
	CLR_LINESPECIAL,
	CLR_LINESPECIALDOUBLE,
	CLR_LINESELECTED,
	CLR_LINEHIGHLIGHT,
	CLR_LINEDRAG,
	CLR_THINGTAG,
	CLR_SECTORTAG,
	CLR_THINGUNKNOWN,
	CLR_THINGSELECTED,
	CLR_THINGHIGHLIGHT,
	CLR_MULTISELECT,
	CLR_GRID,
	CLR_GRID64,
	CLR_LINEBLOCKSOUND,
	CLR_MAPBOUNDARY
};


// Thing images
enum ENUM_THINGIMAGES
{
	TI_ARROW0,
	TI_ARROW45,
	TI_ARROW90,
	TI_ARROW135,
	TI_ARROW180,
	TI_ARROW225,
	TI_ARROW270,
	TI_ARROW315,
	TI_DOT,
	TI_UNKNOWN
};


// Editing modes
enum ENUM_EDITMODE
{
	EM_VERTICES,
	EM_LINES,
	EM_SECTORS,
	EM_THINGS,
	EM_CAMERA
};

// Editing sub modes
enum ENUM_EDITSUBMODE
{
	ESM_NONE,
	ESM_DRAGGING,
	ESM_DRAWING,
	ESM_SELECTING
};

// Standard linedef flags
enum ENUM_LINEDEFFLAGS
{
	LDF_IMPASSIBLE = 1,
	LDF_BLOCKMONSTER = 2,
	LDF_TWOSIDED = 4,
	LDF_UPPERUNPEGGED = 8,
	LDF_LOWERUNPEGGED = 16,
	LDF_SECRET = 32,
	LDF_BLOCKSOUND = 64,
	LDF_HIDDEN = 128,
	LDF_SHOWN = 256
};

// Sector slope flags
enum ENUM_SECTORSLOPEFLAGS
{
	SSF_NONE = 0,
	SSF_FLOOR = 1,
	SSF_CEILING = 2
};


// ZDoom slope line effects
#define LINEEFFECT_SLOPE		181
#define LINEEFFECT_SLOPE_FF	340		// Plane Align Floor at front
#define LINEEFFECT_SLOPE_CF	341		// Plane Align Ceiling at front
#define LINEEFFECT_SLOPE_AF	342		// Plane Align Floor and Ceiling at front
#define LINEEFFECT_SLOPE_FB	343		// Plane Align Floor at back
#define LINEEFFECT_SLOPE_CB	344		// Plane Align Ceiling at back
#define LINEEFFECT_SLOPE_AB	345		// Plane Align Floor and Ceiling at back
#define LINEEFFECT_SLOPE_FBCF	346		// Plane Align Floor at back and Ceiling at front
#define LINEEFFECT_SLOPE_FFCB 347		// Plane Align Floor at front and Ceiling at back


