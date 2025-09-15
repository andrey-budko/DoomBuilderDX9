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

// turn off some stupid warnings
#pragma warning (disable: 4786)					// identifier was truncated to 'number' characters in the debug information

// Includes
#include <string>
#include <set>
#include <string.h>
#include <windows.h>
#include "data.h"

// Namespace
using namespace std;

// Word seperators
//char seperators[] = " \n\r()[],:+-<>=!^%&*{};|/";

// Current settings
string linecomment, commentopen, commentclose, stringopen, stringclose, stringescape, terminator;
set<string> keywords;

// This will hold the parsed rtf until released
char* parsedrtf;


// Script_InitSettings: This sets the settings
//----------------------------------------------------------------------------
void APIENTRY Script_InitSettings(char* c_linecomment, char* c_commentopen, char* c_commentclose,
								  char* c_stringopen, char* c_stringclose, char* c_terminator,
								  char* c_stringescape)
{
	// Set the settings
	linecomment = c_linecomment;
	commentopen = c_commentopen;
	commentclose = c_commentclose;
	stringopen = c_stringopen;
	stringclose = c_stringclose;
	stringescape = c_stringescape;
	terminator = c_terminator;
}


// Script_AddKeyword: Adds a keyword
//----------------------------------------------------------------------------
void APIENTRY Script_AddKeyword(char* keyword)
{
	// Add the keyword
	keywords.insert(keyword);
}


// Script_UnloadKeywords: Cleans up
//----------------------------------------------------------------------------
void APIENTRY Script_UnloadKeywords()
{
	// Release memory
	keywords.clear();
}


// Script_ParseRTF: Parses RTF for keywords and inserts color codes
//----------------------------------------------------------------------------
const char* APIENTRY Script_ParseRTF(char* cleanrtf, char* seperators)
{
	// Lengths
	int cleanrtflen = strlen(cleanrtf);
	int linecommentlen = linecomment.length();
	int commentopenlen = commentopen.length();
	int commentcloselen = commentclose.length();
	int stringopenlen = stringopen.length();
	int stringcloselen = stringclose.length();
	int terminatorlen = terminator.length();
	
	// Codes
	const char color1[] = "\\cf1 ";
	const char color2[] = "\\cf2 ";
	const char color3[] = "\\cf3 ";
	const char color4[] = "\\cf4 ";
	
	// New RTF will be build in this
	string newrtf = " ";
	int lastparsed = 0;
	int lastseperator = 0;
	
	// String search
	char* endptr;
	
	// Lets read the text
	for(int c = 0; c < cleanrtflen; c++)
	{
		// Check if this is for sure NOT a seperator
		if( ((cleanrtf[c] <= 47) || (cleanrtf[c] >= 58)) &&
			((cleanrtf[c] <= 64) || (cleanrtf[c] >= 91)) &&
			((cleanrtf[c] <= 96) || (cleanrtf[c] >= 123)) )
		{
			// Check if this is a seperator
			if((strchr(seperators, cleanrtf[c]) != NULL) && (lastseperator < c))
			{
				// A word has ended here, get the entire word
				char* word = new char[c - lastseperator];
				strncpy(word, cleanrtf + lastseperator + 1, c - lastseperator - 1);
				word[c - lastseperator - 1] = 0;
				
				// Make the word lowercase
				strlwr(word);
				
				// Check if the word is a keyword
				if(keywords.find(word) != keywords.end())
				{
					// Parse up to the keyword, including seperator
					if(lastparsed <= lastseperator) newrtf.append(cleanrtf + lastparsed, lastseperator - lastparsed + 1);
					lastparsed = lastseperator + 1;
					
					// Add color code
					newrtf += color3;
					
					// Parse the keyword, but not the current seperator
					newrtf.append(cleanrtf + lastparsed, c - lastparsed);
					lastparsed = c;
					
					// Add color code
					newrtf += color1;
				}
				
				// Move the last seperator
				lastseperator = c;
				
				// Clean up word
				delete[] word;
			}
			
			// Check for comment open at this position
			if((c < cleanrtflen - commentopenlen) && (strncmp(cleanrtf + c, commentopen.c_str(), commentopenlen) == 0))
			{
				// Parse text before this
				if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed);
				lastparsed = c;
				
				// Add color code
				newrtf += color2;
				
				// Find the next comment close
				endptr = strstr(cleanrtf + lastparsed + commentopenlen, commentclose.c_str());
				if(endptr == NULL) endptr = cleanrtf + cleanrtflen - commentcloselen;
				
				// Move c up to endptr
				c = endptr - cleanrtf;
				
				// Parse text and include this
				if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed + commentcloselen);
				lastparsed = c + commentcloselen;
				lastseperator = lastparsed - 1;
				
				// Add color code
				newrtf += color1;
			}
			// Check for line comment at this position
			else if((c < cleanrtflen - linecommentlen) && (strncmp(cleanrtf + c, linecomment.c_str(), linecommentlen) == 0))
			{
				// Parse text before this
				if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed);
				lastparsed = c;
				
				// Add color code
				newrtf += color2;
				
				// Find the next newline
				endptr = strstr(cleanrtf + lastparsed + linecommentlen, "\n");
				if(endptr == NULL) endptr = cleanrtf + cleanrtflen - 1;
				
				// Move c up to endptr
				c = endptr - cleanrtf;
				
				// Parse text and include this
				if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed + 1);
				lastparsed = c + 1;
				lastseperator = lastparsed;
				
				// Add color code
				newrtf += color1;
			}
			// Check for string open at this position
			else if((c < cleanrtflen - stringopenlen) && (strncmp(cleanrtf + c, stringopen.c_str(), stringopenlen) == 0))
			{
				// Parse text before this
				if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed);
				lastparsed = c;
				
				// Add color code
				newrtf += color4;
				
				do
				{
					// Find the next string close
					endptr = strstr(cleanrtf + c + stringopenlen, stringclose.c_str());
					if(endptr == NULL) endptr = cleanrtf + cleanrtflen - stringcloselen;
					
					// Move c up to endptr
					c = endptr - cleanrtf;
					
					// Check if no stringescape before this
					if(cleanrtf[c - 1] != stringescape.c_str()[0])
					{
						// Parse text and include this
						if(lastparsed < c) newrtf.append(cleanrtf + lastparsed, c - lastparsed + stringcloselen);
						lastparsed = c + stringcloselen;
						lastseperator = lastparsed;
						
						// Add color code
						newrtf += color1;
						
						// Leave the search for the string close
						break;
					}
					
				// Continue until at end of text
				} while(c < cleanrtflen);
			}
		}
	}
	
	// Parse the rest of the text if any left
	if(lastparsed < cleanrtflen) newrtf.append(cleanrtf + lastparsed, cleanrtflen - lastparsed);
	
	// Create standard string
	parsedrtf = new char[newrtf.length() + 0];
	
	// Fill it without first character
	strncpy(parsedrtf, newrtf.c_str() + 1, newrtf.length() - 1);
	parsedrtf[newrtf.length() - 1] = 0;
	
	// Return pointer to new text
	return parsedrtf;
}


// Script_ReleaseRTF: Releases parsed RTF text
//----------------------------------------------------------------------------
void APIENTRY Script_ReleaseRTF()
{
	// Release the text
	delete[] parsedrtf;
}

