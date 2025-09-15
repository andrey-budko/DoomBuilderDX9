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
#include <math.h>
#include "data.h"
#include "math.h"
#include "clip.h"

// turn off some stupid warnings
#pragma warning (disable: 4244)					// conversion from 'type1' to 'type2', possible loss of data

// Clipper structure
struct VCLIPPER
{
	int	used;
	struct VCLIPPER *prev, *next;
	float start, end;
};

// Variables
VCLIPPER	*cliphead;		// First clipper


// InitializeClipper: Initializes the clipper
//-----------------------------------------------------------------------------
void __fastcall InitializeClipper()
{
	// No clippers yet
	cliphead = NULL;
}


// TerminateClipper: Terminates the clipper
//-----------------------------------------------------------------------------
void __fastcall TerminateClipper()
{
	VCLIPPER *node = cliphead;
	VCLIPPER *temp;
	
	// Clean up all clippers
	while (node != NULL)
	{
		// Clean up this clipper
		temp = node;
		node = node->next;
		free(temp);
	}
	
	// All gone
	cliphead = NULL;
}


// TestClipper: Test if the given range is blocked by clippers
//-----------------------------------------------------------------------------
int __fastcall TestClipper(float start, float end)
{
	VCLIPPER *ci;
	ci = cliphead;
	
	// Continue until all clippers tested
	while(ci != NULL)
	{
		// Test clip ranges and return 0 when not visible
		if((start >= ci->start) && (end <= ci->end)) return 0;
		
		// Next clipper
		ci = ci->next;
	}
	
	// Visible!
	return 1;
}


// TestClipRange: Tests if anything between 2 given angels is not clipped.
//-----------------------------------------------------------------------------
int __fastcall TestClipRange(float start, float end)
{
	// Check range direction
	float delta = end - start;
	if(delta < 0) delta += PI2;
	if(delta > PI)
	{
		// Reverse it
		float t = start;
		start = end;
		end = t;
	}
	
	// Check if range wraps around
	if(end < start)
	{
		// Test in two partitions
		return (TestClipper(start, PI2) || TestClipper(0.0f, end));
	}
	else
	{
		// Test given range
		return TestClipper(start, end);
	}
}


// NewClipRange: Creates a new clipper
//-----------------------------------------------------------------------------
VCLIPPER *NewClipRange(float start, float end)
{
	VCLIPPER *node;
	
	// Create clipper
	node = (VCLIPPER*)malloc(sizeof(VCLIPPER));
	
	// Set start and end
	node->start = start;
	node->end = end;
	
	// Return new clipper
	return node;
}


// RemoveClipRange: Removes a clipper
//-----------------------------------------------------------------------------
void RemoveClipRange(VCLIPPER *range)
{
	// Check if first clipper
	if(range == cliphead)
	{
		// Remove first clipper
		cliphead = cliphead->next;
		if(cliphead) cliphead->prev = NULL;
	}
	else
	{
		// Remove clipper
		if(range->prev) range->prev->next = range->next;
		if(range->next) range->next->prev = range->prev;
	}
	
	// Free memory
	free(range);
}


// WriteClipper: Writes a range on the clipbuffer.
//-----------------------------------------------------------------------------
void __fastcall WriteClipper(float start, float end)
{
	VCLIPPER *node, *temp, *prevNode;
	
	// Check if a clipper exists
	if(cliphead != NULL)
	{
		// When range is contained in any old ranges
		// just leave, it has been recorded already
		if(!TestClipRange(start, end)) return;
		
		// Check to see if range contains any old ranges
		node = cliphead;
		while(node != NULL)
		{
			// Old range within this range?
			if(node->start >= start && node->end <= end)
			{
				temp = node;
				node = node->next;
				RemoveClipRange(temp);
			}
			else
			{
				node = node->next;
			}
		}
		
		// Check to see if range overlaps a range (or possibly 2)
		node = cliphead;
		while(node != NULL)
		{
			if(node->start >= start && node->start <= end)
			{
				node->start = start;
				return;
			}
			
			if(node->end >= start && node->end <= end)
			{
				// check for possible merger
				if(node->next)
				{
					// merge two nodes
					if(node->next->start <= end)
					{
						node->end = node->next->end;
						RemoveClipRange(node->next);
					}
					else
					{
						node->end = end;
					}
				}
				else
				{
					node->end = end;
				}
				
				return;
			}
			
			node = node->next;
		}
		
		// Just add range
		node = cliphead;
		prevNode = NULL;
		
		while((node != NULL) && (node->start < end))
		{
			prevNode = node;
			node = node->next;
		}
		
		temp = NewClipRange(start, end);
		
		if(node == NULL)
		{
			temp->next = NULL;
			temp->prev = prevNode;
			if(prevNode) prevNode->next = temp;
			if(cliphead == NULL) cliphead = temp;
		}
		else
		{
			if(node == cliphead)
			{
				temp->next = cliphead;
				cliphead->prev = temp;
				cliphead = temp;
			}
			else
			{
				temp->next = node;
				temp->prev = prevNode;
				if (prevNode) prevNode->next = temp;
				node->prev = temp;
			}
		}
	}
	else
	{
		temp = NewClipRange(start, end);
		temp->next = NULL;
		temp->prev = NULL;
		cliphead = temp;
		return;
	}
}


// WriteClipRange: Writes a range on the clipbuffer.
//-----------------------------------------------------------------------------
void __fastcall WriteClipRange(float start, float end)
{
	// Check range direction
	float delta = end - start;
	if(delta < 0) delta += PI2;
	if(delta > PI)
	{
		// Reverse it
		float t = start;
		start = end;
		end = t;
	}
	
	// Write the range
	if(end < start)
	{
		// Write in two partitions
		WriteClipper(start, PI2);
		WriteClipper(0.0f, end);
	}
	else
	{
		// Write given range
		WriteClipper(start, end);
	}
}


// TestFullClipbuffer: Tests if all angles are clipped out
//-----------------------------------------------------------------------------
int __fastcall TestFullClipbuffer()
{
	// TODO: Check if all angles clipped
	return 0;
}
