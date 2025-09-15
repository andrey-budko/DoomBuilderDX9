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
#define RED(x)					(((x) & 0xFF))
#define GREEN(x)				((((x) >> 8) & 0xFF))
#define BLUE(x)					((((x) >> 16) & 0xFF))
#define CONV_SHIFTS				3
#define CONV_TABLESIZE			32768

// Includes
#include <windows.h>
#include <objbase.h>
#include "const.h"
#include "data.h"
#include "math.h"
#include "renderer.h"

// Color conversion table
unsigned char convtable[CONV_TABLESIZE];


// GetNearestColor: Finds the nearest matching color index from a palette
//----------------------------------------------------------------------------
int __fastcall GetNearestColor(int color, BITMAPRGB* palette, int numcolors)
{
	int difference = 20000000;
	int foundindex = 0;
	int diff;
	int cr = RED(color);
	int cg = GREEN(color);
	int cb = BLUE(color);
	
	// Go for all colors in palette
	for(int c = 0; c < numcolors; c++)
	{
		// Calculate color difference
		diff = abs((int)palette->r - cr) +
			   abs((int)palette->g - cg) +
			   abs((int)palette->b - cb);
		
		// Check if difference is lower
		if(diff < difference)
		{
			// Store index
			foundindex = c;
			difference = diff;
		}
		
		// Next color
		palette++;
	}
	
	// Return found index
	return foundindex;
}


// BuildConversionTable: Makes the color conversion table
//----------------------------------------------------------------------------
void APIENTRY BuildConversionTable(BITMAPRGB* palette, int numcolors)
{
	int r, g, b;
	const int c = 256 >> CONV_SHIFTS;
	
	// Go for all possible colors
	for(r = 0; r < c; r++)
	for(g = 0; g < c; g++)
	for(b = 0; b < c; b++)
	{
		// Assign the nearest palette color to table
		convtable[r + (g << 5) + (b << 10)] = GetNearestColor((r << CONV_SHIFTS) + (g << (8 + CONV_SHIFTS)) + (b << (16 + CONV_SHIFTS)), palette, numcolors);
	}
}


// Flip_Flat: Flips flat data vertically
//----------------------------------------------------------------------------
void APIENTRY Flip_Flat(byte* flat, int tw, int th)
{
	// Calculate half number of rows
	int hr = (int)floor((float)th / 2.0f);
	
	// Allocate memory for enire row
	byte* row = new byte[tw];
	
	// Go for half the number of rows
	for(int r = 0; r < hr; r++)
	{
		// Copy first row to temporary memory
		memcpy(row, flat + r * tw, tw);
		
		// Move second row to first
		//memmove(flat + r * tw, flat + (th - r - 1) * tw, tw);
		memcpy(flat + r * tw, flat + (th - r - 1) * tw, tw);
		
		// Copy first row (from temporary memory) to second
		memcpy(flat + (th - r - 1) * tw, row, tw);
	}
	
	// Destroy memory
	delete[] row;
}


// Draw_Image: Draws from Image data format
//----------------------------------------------------------------------------
void APIENTRY Draw_Image(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, byte transparency_index, byte alternate_index)
{
	byte* dp = NULL;
	int pxls, sr, p;
	int tx, ty;
	
	// Read the patch width and height
	int pw = ((short*)pdata)[0];
	int ph = ((short*)pdata)[1];
	
	// Check if the known length is available
	if(pdatalen >= (8 + pw))
	{
		// Get a pointer to the column offsets
		int* cp = (int*)pdata + 2;
		
		// Go for all columns
		for(int c = 0; c < pw; c++)
		{
			// Check if column pointer is within bounds
			//if(cp[c] <= pdatalen - 2)
			if(cp[c] < pdatalen)
			{
				// Set the data start point for this column
				dp = pdata + cp[c];
				
				// Continue until byte 255 found or end of data
				//while(dp[0] != 255)
				while((dp[0] != 255) && (dp - pdata < pdatalen))
				{
					// Get start row
					sr = dp[0];
					
					// Get number of pixels to draw
					pxls = dp[1];
					
					// Got for all pixels
					for(p = 0; p < pxls; p++)
					{
						// Check if still within bounds
						//if(cp[c] + 2 + p <= pdatalen)
						if(cp[c] + p < pdatalen)
						{
							// Calculate position on texture bitmap
							tx = px + c;
							ty = py + sr + p;
							
							// Check if within bounds
							if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
							{
								// Check if mapped to reserved transparency color
								if(dp[3 + p] == transparency_index)
								{
									// Draw the pixel with alternative index
									texture[(th - ty - 1) * tw + tx] = alternate_index;
								}
								else
								{
									// Draw the pixel
									texture[(th - ty - 1) * tw + tx] = dp[3 + p];
								}
							}
						}
					}
					
					// Move dp to the next row start
					dp += 4 + pxls;
				}
			}
		}
	}
}



// Draw_Flat: Draws from Flat data format
//----------------------------------------------------------------------------
void APIENTRY Draw_Flat(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, int pw, int ph, byte transparency_index, byte alternate_index)
{
	int x, y, tx, ty, fp;
	
	// Check if the known length is available
	if(pdatalen >= (8 + pw))
	{
		// Go for all pixels in x and y
		for(x = 0; x < pw; x++)
		for(y = 0; y < ph; y++)
		{
			// Calculate data position in flat data
			fp = y * pw + x;
			
			// Calculate position on texture bitmap
			tx = px + x;
			ty = py + y;
			
			// Check if within bounds
			if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
			{
				// Check if mapped to reserved transparency color
				if(pdata[fp] == transparency_index)
				{
					// Draw the pixel with alternative index
					texture[(th - ty - 1) * tw + tx] = alternate_index;
				}
				else
				{
					// Draw the pixel
					texture[(th - ty - 1) * tw + tx] = pdata[fp];
				}
			}
		}
	}
}


// Draw_BitmapP8: Draws from 8 bit paletted Bitmap data format
//----------------------------------------------------------------------------
void APIENTRY Draw_BitmapP8(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, int pw, int ph, byte transparency_index, byte alternate_index)
{
	int x, y, tx, ty, fp, ri, gi, bi, palcolor;
	BITMAPFILEHEADER* bitmapheader;
	BITMAPINFOHEADER* bitmapinfo;
	byte* bitmapdata;
	RGBQUAD* palette;
	
	// Get the header
	bitmapheader = (BITMAPFILEHEADER*)pdata;
	
	// Test for bitmap signature
	if( (((char*)bitmapheader)[0] == 'B') &&
		(((char*)bitmapheader)[1] == 'M') )
	{
		// Get the info and palette
		bitmapinfo = (BITMAPINFOHEADER*)(pdata + sizeof(BITMAPFILEHEADER));
		palette = (RGBQUAD*)(pdata + sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER));
		
		// Get the data
		bitmapdata = pdata + bitmapheader->bfOffBits;
		
		// Adjust width and height
		pw = bitmapinfo->biWidth;
		ph = bitmapinfo->biHeight;
		
		// Go for all pixels in x and y
		for(y = 0; y < ph; y++)
		for(x = 0; x < pw; x++)
		{
			// Calculate data position in bitmap data
			fp = y * pw + x;
			
			// Calculate position on texture bitmap
			tx = px + x;
			ty = py + y;
			
			// Check if within bounds
			if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
			{
				// Make color index
				ri = palette[bitmapdata[fp]].rgbRed >> CONV_SHIFTS;
				gi = palette[bitmapdata[fp]].rgbGreen >> CONV_SHIFTS;
				bi = palette[bitmapdata[fp]].rgbBlue >> CONV_SHIFTS;
				
				// Get palette color
				palcolor = convtable[ri + (gi << 5) + (bi << 10)];
				
				// Check if mapped to reserved transparency color
				if(palcolor == transparency_index)
				{
					// Draw the pixel with alternative index
					texture[ty * tw + tx] = alternate_index;
				}
				else
				{
					// Draw the pixel
					texture[ty * tw + tx] = palcolor;
				}
			}
		}
	}
}


// Draw_BitmapB5G6R5: Draws from 16 bit Bitmap data format
//----------------------------------------------------------------------------
void APIENTRY Draw_BitmapB5G6R5(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, int pw, int ph, byte transparency_index, byte alternate_index)
{
	int x, y, tx, ty, fp, ri, gi, bi, palcolor;
	BITMAPFILEHEADER* bitmapheader;
	BITMAPINFOHEADER* bitmapinfo;
	byte* bitmapdata;
	
	// Get the header
	bitmapheader = (BITMAPFILEHEADER*)pdata;
	
	// Test for bitmap signature
	if( (((char*)bitmapheader)[0] == 'B') &&
		(((char*)bitmapheader)[1] == 'M') )
	{
		// Get the info
		bitmapinfo = (BITMAPINFOHEADER*)(pdata + sizeof(BITMAPFILEHEADER));
		
		// Get the data
		bitmapdata = pdata + bitmapheader->bfOffBits;
		
		// Adjust width and height
		pw = bitmapinfo->biWidth;
		ph = bitmapinfo->biHeight;
		
		// Go for all pixels in x and y
		for(y = 0; y < ph; y++)
		for(x = 0; x < pw; x++)
		{
			// Calculate data position in bitmap data
			fp = (y * pw + x) * 2;
			
			// Calculate position on texture bitmap
			tx = px + x;
			ty = py + y;
			
			// Check if within bounds
			if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
			{
				// Make color index
				ri = (*((int*)&bitmapdata[fp]) & 0x1F) >> CONV_SHIFTS;
				gi = (*((int*)&bitmapdata[fp]) >> 5 & 0x3F) >> CONV_SHIFTS;
				bi = (*((int*)&bitmapdata[fp]) >> 11 & 0x1F) >> CONV_SHIFTS;
				
				// Get palette color
				palcolor = convtable[ri + (gi << 5) + (bi << 10)];
				
				// Check if mapped to reserved transparency color
				if(palcolor == transparency_index)
				{
					// Draw the pixel with alternative index
					texture[ty * tw + tx] = alternate_index;
				}
				else
				{
					// Draw the pixel
					texture[ty * tw + tx] = palcolor;
				}
			}
		}
	}
}


// Draw_BitmapB8G8R8: Draws from 24 bit Bitmap data format
//----------------------------------------------------------------------------
void APIENTRY Draw_BitmapB8G8R8(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, int pw, int ph, byte transparency_index, byte alternate_index)
{
	int x, y, tx, ty, fp, ri, gi, bi, palcolor;
	BITMAPFILEHEADER* bitmapheader;
	BITMAPINFOHEADER* bitmapinfo;
	byte* bitmapdata;
	
	// Get the header
	bitmapheader = (BITMAPFILEHEADER*)pdata;
	
	// Test for bitmap signature
	if( (((char*)bitmapheader)[0] == 'B') &&
		(((char*)bitmapheader)[1] == 'M') )
	{
		// Get the info
		bitmapinfo = (BITMAPINFOHEADER*)(pdata + sizeof(BITMAPFILEHEADER));
		
		// Get the data
		bitmapdata = pdata + bitmapheader->bfOffBits;
		
		// Adjust width and height
		pw = bitmapinfo->biWidth;
		ph = bitmapinfo->biHeight;
		
		// Go for all pixels in x and y
		for(y = 0; y < ph; y++)
		for(x = 0; x < pw; x++)
		{
			// Calculate data position in bitmap data
			fp = (y * pw + x) * 3;
			
			// Calculate position on texture bitmap
			tx = px + x;
			ty = py + y;
			
			// Check if within bounds
			if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
			{
				// Make color index
				ri = int(bitmapdata[fp + 2]) >> CONV_SHIFTS;
				gi = int(bitmapdata[fp + 1]) >> CONV_SHIFTS;
				bi = int(bitmapdata[fp + 0]) >> CONV_SHIFTS;
				
				// Get palette color
				palcolor = convtable[ri + (gi << 5) + (bi << 10)];
				
				// Check if mapped to reserved transparency color
				if(palcolor == transparency_index)
				{
					// Draw the pixel with alternative index
					texture[ty * tw + tx] = alternate_index;
				}
				else
				{
					// Draw the pixel
					texture[ty * tw + tx] = palcolor;
				}
			}
		}
	}
}


// Draw_BitmapA8B8G8R8: Draws from 32 bit Bitmap data format
//----------------------------------------------------------------------------
void APIENTRY Draw_BitmapA8B8G8R8(byte* texture, int tw, int th, byte* pdata, int pdatalen, int px, int py, int pw, int ph, byte transparency_index, byte alternate_index)
{
	int x, y, tx, ty, fp, ri, gi, bi, ai, palcolor;
	BITMAPFILEHEADER* bitmapheader;
	BITMAPINFOHEADER* bitmapinfo;
	byte* bitmapdata;
	
	// Get the header
	bitmapheader = (BITMAPFILEHEADER*)pdata;
	
	// Test for bitmap signature
	if( (((char*)bitmapheader)[0] == 'B') &&
		(((char*)bitmapheader)[1] == 'M') )
	{
		// Get the info
		bitmapinfo = (BITMAPINFOHEADER*)(pdata + sizeof(BITMAPFILEHEADER));
		
		// Get the data
		bitmapdata = pdata + bitmapheader->bfOffBits;
		
		// Adjust width and height
		pw = bitmapinfo->biWidth;
		ph = bitmapinfo->biHeight;
		
		// Go for all pixels in x and y
		for(y = 0; y < ph; y++)
		for(x = 0; x < pw; x++)
		{
			// Calculate data position in bitmap data
			fp = (y * pw + x) * 4;
			
			// Calculate position on texture bitmap
			tx = px + x;
			ty = py + y;
			
			// Check if within bounds
			if((tx > -1) && (tx < tw) && (ty > -1) && (ty < th))
			{
				// Check alpha byte
				ai = int(bitmapdata[fp + 0]);
				if(ai > 32)
				{
					// Make color index
					ri = int(bitmapdata[fp + 3]) >> CONV_SHIFTS;
					gi = int(bitmapdata[fp + 2]) >> CONV_SHIFTS;
					bi = int(bitmapdata[fp + 1]) >> CONV_SHIFTS;
					
					// Get palette color
					palcolor = convtable[ri + (gi << 5) + (bi << 10)];
					
					// Check if mapped to reserved transparency color
					if(palcolor == transparency_index)
					{
						// Draw the pixel with alternative index
						texture[ty * tw + tx] = alternate_index;
					}
					else
					{
						// Draw the pixel
						texture[ty * tw + tx] = palcolor;
					}
				}
				else
				{
					// Draw the pixel transparent
					texture[ty * tw + tx] = transparency_index;
				}
			}
		}
	}
}
