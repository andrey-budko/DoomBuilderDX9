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


// Include math library
#include <math.h>

// Definitions
#define MIN(x, y) (x < y ? x : y)
#define MAX(x, y) (x > y ? x : y)
#define MAXQUADRANT			32767
#define MAXQUADRANTf		32767.0f
#define CLIPLINE_LEFT		1
#define CLIPLINE_RIGHT		2
#define CLIPLINE_TOP		4
#define CLIPLINE_BOTTOM		8
#define ENDLESS_DISTANCE	2147483640			// Max int
#define SMALL_FLOAT			0.00000001f
#define PI							3.141592653589793f
#define PI2							6.283185307179586f
#define PIHALF						1.5707963267948965f


//----------------------------------------------------------------------------
int clipline(float &x1, float &y1, float &x2, float &y2, float xmax, float ymax);
bool pointinsidedefs(MAPVERTEX* vertices, MAPLINEDEF* linedefs, MAPSIDEDEF* sidedefs, int vertex, int* sideslist, int numsides);
float __fastcall IntersectLine(FPOINT3D &r1, FPOINT3D &r2, float &x3, float &y3, float &x4, float &y4);
int __fastcall IntersectLine(MAPVERTEX* v1, MAPVERTEX* v2, int x3, int y3, int x4, int y4);
float __fastcall IntersectLineXZ(FPOINT3D &r1, FPOINT3D &r2, float &x3, float &z3, float &x4, float &z4);


// angle: Returns the angle between two points
//----------------------------------------------------------------------------
float inline angle(const int x1, const int y1, const int x2, const int y2)
{
    return atan2(float(y2 - y1), float(x2 - x1));
}


// distance: Returns the distance between two points
//----------------------------------------------------------------------------
float inline distance(const int x1, const int y1, const int x2, const int y2)
{
	int x3 = x2 - x1;
	int y3 = y2 - y1;
	
	if((x3 > MAXQUADRANT) || (y3 > MAXQUADRANT) ||
	   (x3 < -MAXQUADRANT) || (y3 < -MAXQUADRANT))
		return (float)sqrt(double(x3 * x3) + double(y3 * y3));
	else
		return sqrt(float(x3 * x3 + y3 * y3));
}


// distance: Returns the distance between two points
//----------------------------------------------------------------------------
float inline distance(const float x1, const float y1, const float x2, const float y2)
{
	float x3 = x2 - x1;
	float y3 = y2 - y1;
	
	if((x3 > MAXQUADRANTf) || (y3 > MAXQUADRANTf) ||
	   (x3 < -MAXQUADRANTf) || (y3 < -MAXQUADRANTf))
		return (float)sqrt(double(x3 * x3) + double(y3 * y3));
	else
		return sqrt(x3 * x3 + y3 * y3);
}


// point_in_rect: Checks if a point intersects a rectangle
//----------------------------------------------------------------------------
int inline point_in_rect(const int x, const int y, const int left, const int right, const int top, const int bottom)
{
	return (x >= left) && (x <= right) && (y >= top) && (y <= bottom);
}


// point_in_rect: Checks if a point intersects a rectangle
//----------------------------------------------------------------------------
int inline point_in_rect(const int x, const int y, const RECT r)
{
	return (x >= r.left) && (x <= r.right) && (y >= r.top) && (y <= r.bottom);
}


// point_in_rect: Checks if a point intersects a rectangle
//----------------------------------------------------------------------------
int inline point_in_rect(const float x, const float y, const SRECT r)
{
	return (x >= r.left) && (x <= r.right) && (y >= r.top) && (y <= r.bottom);
}


// side_of_line: Calculates on which side of the line a vertex is
// returns < 0 for front (right) side, > 0 for back (left) side and 0 for on the line
//----------------------------------------------------------------------------
float inline side_of_line(int x1, int y1, int x2, int y2, int vx, int vy)
{
	return float(vy - y1) * float(x2 - x1) - float(vx - x1) * float(y2 - y1);
}


// side_of_line: Calculates on which side of the line a vertex is
// returns < 0 for front (right) side, > 0 for back (left) side and 0 for on the line
//----------------------------------------------------------------------------
float inline side_of_line(float x1, float y1, float x2, float y2, int vx, int vy)
{
	return (float(vy) - y1) * (x2 - x1) - (float(vx) - x1) * float(y2 - y1);
}


// side_of_line: Calculates on which side of the line a vertex is
// returns < 0 for front (right) side, > 0 for back (left) side and 0 for on the line
//----------------------------------------------------------------------------
float inline side_of_line(float x1, float y1, float x2, float y2, float vx, float vy)
{
	return (vy - y1) * (x2 - x1) - (vx - x1) * (y2 - y1);
}


// point_near_line: Returns true when a point is in the bounding box of a line
// Usefull for optimization nefore distance_to_line when the max distance is known
//----------------------------------------------------------------------------
int inline point_near_line(float x1, float y1, float x2, float y2, float vx, float vy, int maxdistance)
{
	float rleft, rright, rtop, rbottom;
	
	// Create rect coordinates
	if(x1 > x2)
	{
		rleft = x2 - maxdistance;
		rright = x1 + maxdistance;
	}
	else
	{
		rleft = x1 - maxdistance;
		rright = x2 + maxdistance;
	}
	if(y1 > y2)
	{
		rtop = y2 - maxdistance;
		rbottom = y1 + maxdistance;
	}
	else
	{
		rtop = y1 - maxdistance;
		rbottom = y2 + maxdistance;
	}
	
	// Return true when point is in rect
	return (vx >= rleft) && (vx <= rright) && (vy >= rtop) && (vy <= rbottom);
}


// distance_to_line: Calculates the shortest distance between a vertex and a line
//----------------------------------------------------------------------------
float inline distance_to_line(float x1, float y1, float x2, float y2, float vx, float vy)
{
	float u;
	
	// Get length of linedef
	float ld = distance(x1, y1, x2, y2);
	
	// Check if not zero length
	if(fabs(ld) > SMALL_FLOAT)
	{
		// Get the intersection offset
		u = ((vx - x1) * (x2 - x1) + (vy - y1) * (y2 - y1)) / (ld * ld);
		
		// Limit intersection offset to the line
		float lbound = 1 / ld;
		float ubound = 1 - lbound;
		if(u < lbound) u = lbound;
		if(u > ubound) u = ubound;
	}
	else
	{
		u = 0.0f;
	}
	
	// Calculate intersection point
	float ix = x1 + u * (x2 - x1);
	float iy = y1 + u * (y2 - y1);
	
	// Return distance between intersection and point
	// which is the shortest distance to the line
	return distance(vx, vy, ix, iy);
}


// intercept_vertex: Gives the vertex for two lines intersection
//----------------------------------------------------------------------------
float inline intercept_vertex(FPOINT* start, FPOINT* end, CLIPPER* fdiv, FPOINT* inter)
{
	float ax = start->x, ay = start->y, bx = end->x, by = end->y;
	float cx = fdiv->x, cy = fdiv->y, dx = cx + fdiv->dx, dy = cy + fdiv->dy;
	
	/*
			(YA-YC)(XD-XC)-(XA-XC)(YD-YC)
		r = -----------------------------  (eqn 1)
			(XB-XA)(YD-YC)-(YB-YA)(XD-XC)
	*/
	float r = ((ay - cy) * (dx - cx) - (ax - cx) * (dy - cy)) / ((bx - ax) * (dy - cy) - (by - ay) * (dx - cx));
	
	/*
		XI = XA + r * (XB - XA)
		YI = YA + r * (YB - YA)
	*/
	inter->x = ax + r * (bx - ax);
	inter->y = ay + r * (by - ay);
	return r;
}


// length3: Gives the length between two 3D vertices
//----------------------------------------------------------------------------
float inline length3(FPOINT3D v1, FPOINT3D v2)
{
	FPOINT3D v;
	v.x = v2.x - v1.x;
	v.y = v2.y - v1.y;
	v.z = v2.z - v1.z;
	return sqrt(v.x * v.x + v.y * v.y + v.z * v.z);
}


// normalize3: Gives the normalized 3D vector
//----------------------------------------------------------------------------
FPOINT3D inline normalize3(FPOINT3D v)
{
	FPOINT3D u;
	float norm = sqrt(v.x * v.x + v.y * v.y + v.z * v.z);
	if(fabs(norm) > SMALL_FLOAT)
	{
		float scl = 1.0f / norm;
		u.x = v.x * scl;
		u.y = v.y * scl;
		u.z = v.z * scl;
	}
	else
	{
		u.x = 0.0f;
		u.y = 0.0f;
		u.z = 0.0f;
	}
	return u;
}


// cross3: Gives the cross product of 2 3D vectors
//----------------------------------------------------------------------------
FPOINT3D inline cross3(FPOINT3D v1, FPOINT3D v2)
{
	FPOINT3D v;
	v.x = v1.y * v2.z - v1.z * v2.y;
	v.y = v1.z * v2.x - v1.x * v2.z;
	v.z = v1.x * v2.y - v1.y * v2.x;
	return v;
}


// int_in_array: Returns true when an int is found in an array of ints
//----------------------------------------------------------------------------
bool inline int_in_array(int* array, int numitems, int find)
{
	// Go for all items
	for(int i = 0; i < numitems; i++)
	{
		// Check if this is the item to find
		if(array[i] == find) return true;
	}
	
	// Not found
	return false;
}
