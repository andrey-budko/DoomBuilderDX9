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


#define MAX_SSECTORS_VERTICES		200

struct MAPTHING
{
	int tag;
	int x;
	int y;
	int z;
	int angle;
	int thing;
	int flag;
	int effect;
	int arg[5];
	
	// Optimization variables
	int category;
	int color;
	int image;
	int size;
	int height;
	int hangs;
	int selected;
	int argref[5];
	int sector;
};

struct MAPLINEDEF
{
	int v1;
	int v2;
	int flags;
	int effect;
	int tag;
	int arg[5];
	int s1;
	int s2;
	
	// Optimization variables
	int selected;
	int argref[5];
};

struct MAPSIDEDEF
{
	int tx;
	int ty;
	wchar_t* upper;
	wchar_t* lower;
	wchar_t* middle;
	int sector;
	
	// Optimization variables
	int linedef;
	int middletop;
	int middlebottom;
};

struct MAPVERTEX
{
	float x;
	float y;
	
	// Optimization variables
	int selected;
};

struct MAPSECTOR
{
	int hfloor;
	int hceiling;
	wchar_t* tfloor;
	wchar_t* tceiling;
	int brightness;
	int special;
	int tag;
	
	// Optimization variables
	int selected;
	int visible;
};

struct MAPSPLIT
{
	int x;
	int y;
	int dx;
	int dy;
};

struct MAPSEG
{
	int v1;
	int v2;
	int angle;
	int linedef;
	int side;
	int offset;
};

struct MAPSSECTOR
{
	int startseg;
	int numsegs;
	
	// Optimization variables
	int sector;
	int numvertices;
	POINT vertices[MAX_SSECTORS_VERTICES];
};

struct MAPNODE
{
	int x;
	int y;
	int dx;
	int dy;
	
	int rtop;
	int rbottom;
	int rleft;
	int rright;
	
	int ltop;
	int lbottom;
	int lleft;
	int lright;
	
	int right;
	int left;
};

struct D3DVERTEX
{
	float x, y, z;
	float nx, ny, nz;
	float tu, tv;
};

struct CLIPPER
{
	float x;
	float y;
	float dx;
	float dy;
};

struct FPOINT
{
	float x;
	float y;
};

struct FPOINT3D
{
	float x;
	float y;
	float z;
};

struct CHARDESC
{
	float u1;
	float v1;
	float u2;
	float v2;
	float width;
	float height;
};

struct SRECT
{
	float left;
	float right;
	float top;
	float bottom;
};

struct FLATVERTEX
{
	float x;
	float y;
	float z;
	float rhw;
	long diffuse;
	float tu1;
	float tv1;
};

struct BITMAPRGB
{
	unsigned char b;
	unsigned char g;
	unsigned char r;
	unsigned char x;
};

struct THINGFILTERS
{
	int filtermode;
	int category;
	int flags;
};


int __fastcall NearestLinedefF(int x, int y, MAPVERTEX* vertices, MAPLINEDEF* linedefs, int numlinedefs, int &dist);
int __fastcall ThingFiltered(MAPTHING* thing, int filterthings, THINGFILTERS* filter);
void __fastcall ReallocIntP(int* &intarray, int oldlength, int newlength);
int FurthestSectorVertexF(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int numlinedefs, int line, int sector);

