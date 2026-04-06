#pragma once

#ifndef _TOOLBARCONFIGS_H
#define _TOOLBARCONFIGS_H

#include "resource.h"
#include <commctrl.h>

// ================================================================================================
// NT4 toolbar bitmap glyph indices.
//
// The toolbar uses two NT4 comctl32 bitmap strips loaded sequentially into
// a single image list:
//   STD strip  (15 glyphs, indices 0-14):  standard comctl32 icons
//   VIEW strip (12 glyphs, indices 15-26): view/folder comctl32 icons
//

// STD strip glyph indices (matching comctl32 STD_* constants)
#define GLYPHIDX_CUT                    0   // STD_CUT
#define GLYPHIDX_COPY                   1   // STD_COPY
#define GLYPHIDX_PASTE                  2   // STD_PASTE
#define GLYPHIDX_UNDO                   3   // STD_UNDO
// index 4 = STD_REDO (unused)
#define GLYPHIDX_DELETE                 5   // STD_DELETE
// indices 6-9 unused (filenew, fileopen, filesave, printpre)
#define GLYPHIDX_PROPERTIES             10  // STD_PROPERTIES

// VIEW strip glyph indices (offset by NUM_STD_GLYPHS)
#define NUM_STD_GLYPHS                  15
#define NUM_VIEW_GLYPHS                 12
#define VIEW_STRIP_OFFSET               NUM_STD_GLYPHS

#define GLYPHIDX_LARGEICONS             (VIEW_STRIP_OFFSET + 0)   // VIEW_LARGEICONS
#define GLYPHIDX_SMALLICONS             (VIEW_STRIP_OFFSET + 1)   // VIEW_SMALLICONS
#define GLYPHIDX_LIST                   (VIEW_STRIP_OFFSET + 2)   // VIEW_LIST
#define GLYPHIDX_DETAILS                (VIEW_STRIP_OFFSET + 3)   // VIEW_DETAILS
// indices 19-22 unused (sort by name/size/date/type)
#define GLYPHIDX_UPONELEVEL             (VIEW_STRIP_OFFSET + 8)   // VIEW_PARENTFOLDER
#define GLYPHIDX_MAPNETDRIVE            (VIEW_STRIP_OFFSET + 9)   // VIEW_NETCONNECT
#define GLYPHIDX_DISCONNECTNETDRIVE     (VIEW_STRIP_OFFSET + 10)  // VIEW_NETDISCONNECT

#define NUM_TOOLBAR_GLYPHS              (NUM_STD_GLYPHS + NUM_VIEW_GLYPHS)

// ================================================================================================
// Toolbar button definition structure
//

struct TB_BUTTON_DEF
{
	int     idCommand;
	int     iBitmap;
	BYTE    fsStyle;
	int     idsString;      // string resource ID for tooltip
};

// ================================================================================================
// NT4 button catalog — fixed set matching Windows NT 4.0 Explorer toolbar.
//

static const TB_BUTTON_DEF c_tbAllButtons_NT4[] =
{
	{ TBIDM_PREVIOUSFOLDER,  GLYPHIDX_UPONELEVEL,            BTNS_BUTTON,      IDS_TB_UP                 },
	{ TBIDM_CONNECT,         GLYPHIDX_MAPNETDRIVE,           BTNS_BUTTON,      IDS_TB_MAPNETDRIVE        },
	{ TBIDM_DISCONNECT,      GLYPHIDX_DISCONNECTNETDRIVE,    BTNS_BUTTON,      IDS_TB_DISCONNECTNETDRIVE },
	{ TBIDM_CUT,             GLYPHIDX_CUT,                   BTNS_BUTTON,      IDS_TB_CUT                },
	{ TBIDM_COPY,            GLYPHIDX_COPY,                  BTNS_BUTTON,      IDS_TB_COPY               },
	{ TBIDM_PASTE,           GLYPHIDX_PASTE,                 BTNS_BUTTON,      IDS_TB_PASTE              },
	{ TBIDM_UNDO,            GLYPHIDX_UNDO,                  BTNS_BUTTON,      IDS_TB_UNDO               },
	{ TBIDM_DELETE,          GLYPHIDX_DELETE,                 BTNS_BUTTON,      IDS_TB_DELETE             },
	{ TBIDM_PROPERTIES,      GLYPHIDX_PROPERTIES,            BTNS_BUTTON,      IDS_TB_PROPERTIES         },
	{ TBIDM_VIEW_LARGEICONS, GLYPHIDX_LARGEICONS,            BTNS_CHECKGROUP,  IDS_TB_LARGEICONS         },
	{ TBIDM_VIEW_SMALLICONS, GLYPHIDX_SMALLICONS,            BTNS_CHECKGROUP,  IDS_TB_SMALLICONS         },
	{ TBIDM_VIEW_LIST,       GLYPHIDX_LIST,                  BTNS_CHECKGROUP,  IDS_TB_LIST               },
	{ TBIDM_VIEW_DETAILS,    GLYPHIDX_DETAILS,               BTNS_CHECKGROUP,  IDS_TB_DETAILS            },
};

static const int c_nAllButtons_NT4 = ARRAYSIZE(c_tbAllButtons_NT4);

// NT4 default button layout: array of command IDs (0 = separator).
static const int c_defaultLayout_NT4[] =
{
	TBIDM_PREVIOUSFOLDER,
	0,                          // separator
	TBIDM_CONNECT,
	TBIDM_DISCONNECT,
	0,                          // separator
	TBIDM_CUT,
	TBIDM_COPY,
	TBIDM_PASTE,
	TBIDM_UNDO,
	0,                          // separator
	TBIDM_DELETE,
	TBIDM_PROPERTIES,
	0,                          // separator
	TBIDM_VIEW_LARGEICONS,
	TBIDM_VIEW_SMALLICONS,
	TBIDM_VIEW_LIST,
	TBIDM_VIEW_DETAILS,
};

static const int c_nDefaultLayout_NT4 = ARRAYSIZE(c_defaultLayout_NT4);

// ================================================================================================
// Bitmap ID accessors
//

struct TB_SKIN_BITMAPS
{
	UINT    idStd;          // STD strip bitmap resource ID
	UINT    idView;         // VIEW strip bitmap resource ID
	int     cx;             // glyph width in pixels
};

inline TB_SKIN_BITMAPS GetActiveBitmapIDs(bool fSmallIcons)
{
	TB_SKIN_BITMAPS bmp = {};
	if (fSmallIcons)
	{
		bmp.idStd  = IDB_NT4_STD_SMALL;
		bmp.idView = IDB_NT4_VIEW_SMALL;
		bmp.cx     = 16;
	}
	else
	{
		bmp.idStd  = IDB_NT4_STD_LARGE;
		bmp.idView = IDB_NT4_VIEW_LARGE;
		bmp.cx     = 24;
	}
	return bmp;
}

// Find a button definition in the catalog by command ID
inline int FindButtonCatalogIndex(int idCommand)
{
	for (int i = 0; i < c_nAllButtons_NT4; i++)
	{
		if (c_tbAllButtons_NT4[i].idCommand == idCommand)
			return i;
	}
	return -1;
}

#endif // _TOOLBARCONFIGS_H
