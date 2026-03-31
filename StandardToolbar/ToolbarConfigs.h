#pragma once

#ifndef _TOOLBARCONFIGS_H
#define _TOOLBARCONFIGS_H

#include "resource.h"
#include "../util/util.h"
#include <commctrl.h>

// ================================================================================================
// Bitmap glyph indices in the shell toolbar bitmaps.
// Both XP and 2K bitmap strips are pre-arranged to use this same ordering.
// (XP: from Server 2003 browseui shdef.bmp / shell32 tb_sh_def.bmp layout)
// (2K: rearranged from win2k browseui tbdef.bmp + shdef.bmp to match)
//

#define GLYPHIDX_BACK                   0
#define GLYPHIDX_FORWARD                1
#define GLYPHIDX_UPONELEVEL             2
#define GLYPHIDX_SEARCH                 3
#define GLYPHIDX_FOLDERS                4
#define GLYPHIDX_MOVETO                 5
#define GLYPHIDX_COPYTO                 6
#define GLYPHIDX_DELETE                 7
#define GLYPHIDX_UNDO                   8
#define GLYPHIDX_VIEWS                  9
#define GLYPHIDX_STOP                   10
#define GLYPHIDX_REFRESH                11
#define GLYPHIDX_HOME                   12
#define GLYPHIDX_MAPNETDRIVE            13
#define GLYPHIDX_DISCONNECTNETDRIVE     14
#define GLYPHIDX_FAVORITES              15
#define GLYPHIDX_HISTORY                16
#define GLYPHIDX_FULLSCREEN             17

// Indices 18-22 are appended at runtime from shell32 system bitmaps
#define GLYPHIDX_PROPERTIES             18
#define GLYPHIDX_CUT                    19
#define GLYPHIDX_COPY                   20
#define GLYPHIDX_PASTE                  21
#define GLYPHIDX_FOLDEROPTIONS          22

#define NUM_TOOLBAR_GLYPHS              23

// ================================================================================================
// Toolbar button definition structure (shared between skins)
//

struct TB_BUTTON_DEF
{
	int     idCommand;
	int     iBitmap;
	BYTE    fsStyle;
	int     idsString;      // string resource ID for tooltip/text
	BOOL    bEnabled;       // initial enabled state
};

// Maximum number of buttons in a saved toolbar layout
#define MAX_SAVED_BUTTONS 32

// ================================================================================================
// Skin configuration: returned by accessor functions
//

struct TB_SKIN_BITMAPS
{
	UINT    idDef;          // default state bitmap resource ID
	UINT    idHot;          // hot (hover) state bitmap resource ID
	int     cx;             // glyph width in pixels
};

struct TB_SKIN_DEFAULTS
{
	DWORD   smallIcons;     // 0 = large, 1 = small
	DWORD   textLabelMode;  // CE_TEXTMODE_* value
};

// ================================================================================================
// XP button catalog
// Full catalog of available toolbar buttons for XP skin.
// Order: browseui buttons (c_tbExplorer) then defview buttons (c_tbDefView).
//

static const TB_BUTTON_DEF c_tbAllButtons_XP[] =
{
	// --- browseui buttons (from Server 2003 itbar.cpp c_tbExplorer[]) ---
	{ TBIDM_BACK,           GLYPHIDX_BACK,              BTNS_DROPDOWN | BTNS_SHOWTEXT, IDS_TB_BACK,              TRUE },
	{ TBIDM_FORWARD,        GLYPHIDX_FORWARD,           BTNS_DROPDOWN,                 IDS_TB_FORWARD,           TRUE },
	{ TBIDM_STOPDOWNLOAD,   GLYPHIDX_STOP,              BTNS_BUTTON,                   IDS_TB_STOP,              TRUE },
	{ TBIDM_REFRESH,        GLYPHIDX_REFRESH,           BTNS_BUTTON,                   IDS_TB_REFRESH,           TRUE },
	{ TBIDM_HOME,           GLYPHIDX_HOME,              BTNS_BUTTON,                   IDS_TB_HOME,              TRUE },
	{ TBIDM_PREVIOUSFOLDER, GLYPHIDX_UPONELEVEL,        BTNS_BUTTON,                   IDS_TB_UP,                TRUE },
	{ TBIDM_CONNECT,        GLYPHIDX_MAPNETDRIVE,       BTNS_BUTTON,                   IDS_TB_MAPNETDRIVE,       TRUE },
	{ TBIDM_DISCONNECT,     GLYPHIDX_DISCONNECTNETDRIVE,BTNS_BUTTON,                   IDS_TB_DISCONNECTNETDRIVE,TRUE },
	{ TBIDM_SEARCH,         GLYPHIDX_SEARCH,            BTNS_BUTTON | BTNS_SHOWTEXT,   IDS_TB_SEARCH,            TRUE },
	{ TBIDM_ALLFOLDERS,     GLYPHIDX_FOLDERS,           BTNS_CHECK | BTNS_SHOWTEXT,    IDS_TB_FOLDERS,           TRUE },
	{ TBIDM_FAVORITES,      GLYPHIDX_FAVORITES,         BTNS_BUTTON | BTNS_SHOWTEXT,   IDS_TB_FAVORITES,         TRUE },
	{ TBIDM_HISTORY,        GLYPHIDX_HISTORY,           BTNS_BUTTON | BTNS_SHOWTEXT,   IDS_TB_HISTORY,           TRUE },
	{ TBIDM_THEATER,        GLYPHIDX_FULLSCREEN,        BTNS_BUTTON,                   IDS_TB_FULLSCREEN,        TRUE },
	// --- defview buttons (from defview.cpp c_tbDefView[]) ---
	{ TBIDM_MOVETO,         GLYPHIDX_MOVETO,            BTNS_BUTTON,                   IDS_TB_MOVETO,            TRUE },
	{ TBIDM_COPYTO,         GLYPHIDX_COPYTO,            BTNS_BUTTON,                   IDS_TB_COPYTO,            TRUE },
	{ TBIDM_DELETE,         GLYPHIDX_DELETE,             BTNS_BUTTON,                   IDS_TB_DELETE,            TRUE },
	{ TBIDM_UNDO,           GLYPHIDX_UNDO,              BTNS_BUTTON,                   IDS_TB_UNDO,              TRUE },
	{ TBIDM_VIEWMENU,       GLYPHIDX_VIEWS,             BTNS_WHOLEDROPDOWN,            IDS_TB_VIEWS,             TRUE },
	{ TBIDM_PROPERTIES,     GLYPHIDX_PROPERTIES,        BTNS_BUTTON,                   IDS_TB_PROPERTIES,        TRUE },
	{ TBIDM_CUT,            GLYPHIDX_CUT,               BTNS_BUTTON,                   IDS_TB_CUT,              TRUE },
	{ TBIDM_COPY,           GLYPHIDX_COPY,              BTNS_BUTTON,                   IDS_TB_COPY,             TRUE },
	{ TBIDM_PASTE,          GLYPHIDX_PASTE,             BTNS_BUTTON,                   IDS_TB_PASTE,            TRUE },
	{ TBIDM_FOLDEROPTIONS,  GLYPHIDX_FOLDEROPTIONS,     BTNS_BUTTON,                   IDS_TB_FOLDEROPTIONS,    TRUE },
};

static const int c_nAllButtons_XP = ARRAYSIZE(c_tbAllButtons_XP);

// XP default button layout: array of command IDs (0 = separator).
static const int c_defaultLayout_XP[] =
{
	TBIDM_BACK,
	TBIDM_FORWARD,
	TBIDM_PREVIOUSFOLDER,
	0,                      // separator
	TBIDM_SEARCH,
	TBIDM_ALLFOLDERS,
	0,                      // separator
	TBIDM_VIEWMENU,
};

static const int c_nDefaultLayout_XP = ARRAYSIZE(c_defaultLayout_XP);

// ================================================================================================
// 2K button catalog
// Full catalog of available toolbar buttons for Windows 2000 skin.
// Same buttons as XP — the glyph indices are shared because the 2K bitmap strips
// are pre-rearranged to match the GLYPHIDX_* ordering.
//
// Style differences from XP:
// - Back: BTNS_DROPDOWN (with text, same as XP)
// - Folders: BTNS_CHECK (check button, same as XP)
// - History: no BTNS_SHOWTEXT by default (IE6 on 2000)
//   (IE5.5 mode adds BTNS_SHOWTEXT to History at runtime)
//

static const TB_BUTTON_DEF c_tbAllButtons_2K[] =
{
	// --- browseui buttons ---
	{ TBIDM_BACK,           GLYPHIDX_BACK,              BTNS_DROPDOWN | BTNS_SHOWTEXT, IDS_TB_BACK,              TRUE },
	{ TBIDM_FORWARD,        GLYPHIDX_FORWARD,           BTNS_DROPDOWN,                 IDS_TB_FORWARD,           TRUE },
	{ TBIDM_STOPDOWNLOAD,   GLYPHIDX_STOP,              BTNS_BUTTON,                   IDS_TB_STOP,              TRUE },
	{ TBIDM_REFRESH,        GLYPHIDX_REFRESH,           BTNS_BUTTON,                   IDS_TB_REFRESH,           TRUE },
	{ TBIDM_HOME,           GLYPHIDX_HOME,              BTNS_BUTTON,                   IDS_TB_HOME,              TRUE },
	{ TBIDM_PREVIOUSFOLDER, GLYPHIDX_UPONELEVEL,        BTNS_BUTTON,                   IDS_TB_UP,                TRUE },
	{ TBIDM_CONNECT,        GLYPHIDX_MAPNETDRIVE,       BTNS_BUTTON,                   IDS_TB_MAPNETDRIVE,       TRUE },
	{ TBIDM_DISCONNECT,     GLYPHIDX_DISCONNECTNETDRIVE,BTNS_BUTTON,                   IDS_TB_DISCONNECTNETDRIVE,TRUE },
	{ TBIDM_SEARCH,         GLYPHIDX_SEARCH,            BTNS_BUTTON | BTNS_SHOWTEXT,   IDS_TB_SEARCH,            TRUE },
	{ TBIDM_ALLFOLDERS,     GLYPHIDX_FOLDERS,           BTNS_CHECK | BTNS_SHOWTEXT,    IDS_TB_FOLDERS,           TRUE },
	{ TBIDM_FAVORITES,      GLYPHIDX_FAVORITES,         BTNS_BUTTON | BTNS_SHOWTEXT,   IDS_TB_FAVORITES,         TRUE },
	{ TBIDM_HISTORY,        GLYPHIDX_HISTORY,           BTNS_BUTTON,                   IDS_TB_HISTORY,           TRUE },
	{ TBIDM_THEATER,        GLYPHIDX_FULLSCREEN,        BTNS_BUTTON,                   IDS_TB_FULLSCREEN,        TRUE },
	// --- defview buttons ---
	{ TBIDM_MOVETO,         GLYPHIDX_MOVETO,            BTNS_BUTTON,                   IDS_TB_MOVETO,            TRUE },
	{ TBIDM_COPYTO,         GLYPHIDX_COPYTO,            BTNS_BUTTON,                   IDS_TB_COPYTO,            TRUE },
	{ TBIDM_DELETE,         GLYPHIDX_DELETE,             BTNS_BUTTON,                   IDS_TB_DELETE,            TRUE },
	{ TBIDM_UNDO,           GLYPHIDX_UNDO,              BTNS_BUTTON,                   IDS_TB_UNDO,              TRUE },
	{ TBIDM_VIEWMENU,       GLYPHIDX_VIEWS,             BTNS_WHOLEDROPDOWN,            IDS_TB_VIEWS,             TRUE },
	{ TBIDM_PROPERTIES,     GLYPHIDX_PROPERTIES,        BTNS_BUTTON,                   IDS_TB_PROPERTIES,        TRUE },
	{ TBIDM_CUT,            GLYPHIDX_CUT,               BTNS_BUTTON,                   IDS_TB_CUT,              TRUE },
	{ TBIDM_COPY,           GLYPHIDX_COPY,              BTNS_BUTTON,                   IDS_TB_COPY,             TRUE },
	{ TBIDM_PASTE,          GLYPHIDX_PASTE,             BTNS_BUTTON,                   IDS_TB_PASTE,            TRUE },
	{ TBIDM_FOLDEROPTIONS,  GLYPHIDX_FOLDEROPTIONS,     BTNS_BUTTON,                   IDS_TB_FOLDEROPTIONS,    TRUE },
};

static const int c_nAllButtons_2K = ARRAYSIZE(c_tbAllButtons_2K);

// 2K default button layout (post-IE6 Windows 2000):
// Back, Forward, Up, sep, Search, Folders, History, sep, MoveTo, CopyTo, Delete, Undo, sep, Views
static const int c_defaultLayout_2K[] =
{
	TBIDM_BACK,
	TBIDM_FORWARD,
	TBIDM_PREVIOUSFOLDER,
	0,                      // separator
	TBIDM_SEARCH,
	TBIDM_ALLFOLDERS,
	TBIDM_HISTORY,
	0,                      // separator
	TBIDM_MOVETO,
	TBIDM_COPYTO,
	TBIDM_DELETE,
	TBIDM_UNDO,
	0,                      // separator
	TBIDM_VIEWMENU,
};

static const int c_nDefaultLayout_2K = ARRAYSIZE(c_defaultLayout_2K);

// ================================================================================================
// Config accessor functions — return the active skin's data based on theme
//

inline const TB_BUTTON_DEF* GetActiveButtonCatalog(ClassicExplorerTheme theme, int* pCount)
{
	if (theme == CLASSIC_EXPLORER_2K)
	{
		if (pCount) *pCount = c_nAllButtons_2K;
		return c_tbAllButtons_2K;
	}
	if (pCount) *pCount = c_nAllButtons_XP;
	return c_tbAllButtons_XP;
}

inline const int* GetActiveDefaultLayout(ClassicExplorerTheme theme, int* pCount)
{
	if (theme == CLASSIC_EXPLORER_2K)
	{
		if (pCount) *pCount = c_nDefaultLayout_2K;
		return c_defaultLayout_2K;
	}
	if (pCount) *pCount = c_nDefaultLayout_XP;
	return c_defaultLayout_XP;
}

inline TB_SKIN_BITMAPS GetActiveBitmapIDs(ClassicExplorerTheme theme, bool fSmallIcons, bool fIE55 = false)
{
	TB_SKIN_BITMAPS bmp = {};
	if (theme == CLASSIC_EXPLORER_2K)
	{
		if (fIE55)
		{
			if (fSmallIcons)
			{
				bmp.idDef = IDB_TB_2K_IE55_SHELL_DEF_16;
				bmp.idHot = IDB_TB_2K_IE55_SHELL_HOT_16;
				bmp.cx = 16;
			}
			else
			{
				bmp.idDef = IDB_TB_2K_IE55_SHELL_DEF_20;
				bmp.idHot = IDB_TB_2K_IE55_SHELL_HOT_20;
				bmp.cx = 20;
			}
		}
		else
		{
			if (fSmallIcons)
			{
				bmp.idDef = IDB_TB_2K_SHELL_DEF_16;
				bmp.idHot = IDB_TB_2K_SHELL_HOT_16;
				bmp.cx = 16;
			}
			else
			{
				bmp.idDef = IDB_TB_2K_SHELL_DEF_20;
				bmp.idHot = IDB_TB_2K_SHELL_HOT_20;
				bmp.cx = 20;  // authentic Win2K "large" icon size
			}
		}
	}
	else
	{
		if (fSmallIcons)
		{
			bmp.idDef = IDB_TB_SHELL_DEF_16;
			bmp.idHot = IDB_TB_SHELL_HOT_16;
			bmp.cx = 16;
		}
		else
		{
			bmp.idDef = IDB_TB_SHELL_DEF_24;
			bmp.idHot = IDB_TB_SHELL_HOT_24;
			bmp.cx = 24;
		}
	}
	return bmp;
}

inline TB_SKIN_DEFAULTS GetDefaultSettings(ClassicExplorerTheme theme)
{
	TB_SKIN_DEFAULTS def = {};
	if (theme == CLASSIC_EXPLORER_2K)
	{
		def.smallIcons = 1;                    // 2K defaults to small icons
		def.textLabelMode = CE_TEXTMODE_SELECTIVE;
	}
	else
	{
		def.smallIcons = 0;                    // XP defaults to large icons
		def.textLabelMode = CE_TEXTMODE_SELECTIVE;
	}
	return def;
}

// Find a button definition in a catalog by command ID
inline int FindButtonCatalogIndex(const TB_BUTTON_DEF* pCatalog, int nCatalog, int idCommand)
{
	for (int i = 0; i < nCatalog; i++)
	{
		if (pCatalog[i].idCommand == idCommand)
			return i;
	}
	return -1;
}

#endif // _TOOLBARCONFIGS_H
