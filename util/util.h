#pragma once
#ifndef _UTIL_H
#define _UTIL_H

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
//#include "ClassicExplorer_i.h"
#include "dllmain.h"

enum ClassicExplorerTheme
{
	CLASSIC_EXPLORER_NONE = -1,
	CLASSIC_EXPLORER_2K = 0,
	CLASSIC_EXPLORER_XP = 1,
};

// Toolbar text label modes (matching XP's IDS_TEXTLABELS/IDS_PARTIALTEXT/IDS_NOTEXTLABELS)
#define CE_TEXTMODE_SHOWTEXT    0   // Show text labels (below icons, no TBSTYLE_LIST)
#define CE_TEXTMODE_SELECTIVE   1   // Selective text on right (TBSTYLE_LIST + MIXEDBUTTONS)
#define CE_TEXTMODE_NOTEXT      2   // No text labels (TB_SETMAXTEXTROWS=0)

// Registered window message for notifying toolbar bands of settings changes
#define CE_WM_SETTINGS_CHANGED_NAME L"ClassicExplorer_SettingsChanged"

// Registered window message for requesting the toolbar customize dialog
#define CE_WM_CUSTOMIZE_TOOLBAR_NAME L"ClassicExplorer_CustomizeToolbar"

namespace CEUtil
{
	struct CESettings
	{
		ClassicExplorerTheme theme = CLASSIC_EXPLORER_NONE;
		DWORD showGoButton = -1;
		DWORD showAddressLabel = -1;
		DWORD showFullAddress = -1;
		DWORD smallIcons = -1;        // 0 = large (24px, default), 1 = small (16px)
		DWORD textLabelMode = -1;     // CE_TEXTMODE_* values, default = CE_TEXTMODE_SELECTIVE
		DWORD ie55Style = -1;         // 1 = IE 5.5 style (2K only: History shows text label)
		DWORD win98Views = -1;        // 1 = 98-style split Views button (click cycles, arrow drops down)

		CESettings(ClassicExplorerTheme t, int a, int b, int f,
		           int icons = -1, int textMode = -1, int ie55 = -1, int w98views = -1)
		{
			theme = t;
			showGoButton = a;
			showAddressLabel = b;
			showFullAddress = f;
			smallIcons = icons;
			textLabelMode = textMode;
			ie55Style = ie55;
			win98Views = w98views;
		}
	};
	CESettings GetCESettings();
	void WriteCESettings(CESettings& toWrite);
	HRESULT GetCurrentFolderPidl(CComPtr<IShellBrowser> pShellBrowser, PIDLIST_ABSOLUTE *pidlOut);
	HRESULT FixExplorerSizes(HWND explorerChild);
	HRESULT FixExplorerSizesIfNecessary(HWND explorerChild);
}

#endif