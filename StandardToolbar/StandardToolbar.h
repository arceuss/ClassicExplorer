#pragma once

#ifndef _STANDARDTOOLBAR_H
#define _STANDARDTOOLBAR_H

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"

#include <commctrl.h>
#include <commoncontrols.h>
#include <string>
#include "../AddressBar/AddressBar.h"

// ================================================================================================
// Standard Buttons toolbar command IDs (from Windows Server 2003 itbdrop.h)
//
#define TBIDM_BACK              0x120
#define TBIDM_FORWARD           0x121
#define TBIDM_HOME              0x122
#define TBIDM_SEARCH            0x123
#define TBIDM_STOPDOWNLOAD      0x124
#define TBIDM_REFRESH           0x125
#define TBIDM_FAVORITES         0x126
#define TBIDM_THEATER           0x128
#define TBIDM_HISTORY           0x12E
#define TBIDM_PREVIOUSFOLDER    0x130
#define TBIDM_CONNECT           0x131
#define TBIDM_DISCONNECT        0x132
#define TBIDM_ALLFOLDERS        0x133

// Our own IDs for DefView-originated buttons (offset from 0x200 to avoid collision)
#define TBIDM_DELETE            0x203
#define TBIDM_UNDO              0x204
#define TBIDM_PROPERTIES        0x205
#define TBIDM_CUT               0x206
#define TBIDM_COPY              0x207
#define TBIDM_PASTE             0x208

// NT4 view mode buttons (BTNS_CHECKGROUP)
#define TBIDM_VIEW_LARGEICONS   0x210
#define TBIDM_VIEW_SMALLICONS   0x211
#define TBIDM_VIEW_LIST         0x212
#define TBIDM_VIEW_DETAILS      0x213

// Per-skin button catalogs, default layouts, bitmap IDs, and accessor functions
#include "ToolbarConfigs.h"

// ================================================================================================
// CStandardToolbar: Implements the Standard Buttons toolbar band (XP / 2K skins).
//

class ATL_NO_VTABLE CStandardToolbar :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CStandardToolbar, &CLSID_CStandardToolbar>,
	public IObjectWithSiteImpl<CStandardToolbar>,
	public IDeskBand2,
	public IDispEventImpl<1, CStandardToolbar, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>,
	public IInputObject
{
private:
	// COM site and shell objects:
	CComPtr<IWebBrowser2>    m_pWebBrowser;
	CComPtr<IShellBrowser>   m_pShellBrowser;
	CComPtr<IPropertyBag>    m_pBrowserBag;
	IInputObjectSite*        m_pSite = nullptr;
	HWND                     m_hWndParent = nullptr;

	// Toolbar control:
	HWND                     m_hWndBand = nullptr;       // container window (band child)
	HWND                     m_hWndToolbar = nullptr;
	HIMAGELIST               m_hilDefault = nullptr;

	// Embedded address bar (NT4: combobox is part of this band):
	CAddressBar              m_addressBar;
	int                      m_comboWidth = 0;           // NT4: LOGPIXELSY * 2 (computed at create time)

	// State:
	bool                     m_bShow = false;
	BOOL                     m_bCanGoBack = FALSE;
	BOOL                     m_bCanGoForward = FALSE;

	// Button string pool:
	int                      m_iStringPool[NUM_TOOLBAR_GLYPHS];

	// Internal helpers:
	HRESULT CreateToolbarWindow(HWND hWndParent);
	void    InitToolbarButtons();
	void    PreloadAllStrings();
	void    UpdateButtonStates();
	void    DestroyToolbarWindow();
	void    UpdateViewModeChecks();

	// Command handlers:
	void    OnBack();
	void    OnForward();
	void    OnUpOneLevel();
	void    OnSearch();
	void    OnFolders();
	void    OnMoveTo();
	void    OnCopyTo();
	void    OnDelete();
	void    OnUndo();
	void    OnStop();
	void    OnRefresh();
	void    OnHome();
	void    OnMapNetDrive();
	void    OnDisconnectNetDrive();
	void    OnFavorites();
	void    OnHistory();
	void    OnFullScreen();
	void    OnProperties();
	void    OnCut();
	void    OnCopy();
	void    OnPaste();
	void    OnFolderOptions();

	// Rebar subclass for command routing:
	void    InstallRebarHook();

	// Window command routing helpers (from Open-Shell patterns):
	void    SendShellTabCommand(int command);
	void    SendDefViewCommand(int command);
	void    SetClassicViewMode(FOLDERVIEWMODE fvm);
	void    GetBrowserBag();

public:
	// Accessors used by the rebar subclass proc:
	HWND    GetToolbarHwnd() const;
	void    HandleCommand(int idCmd);
	void    ShowBackForwardMenu(int idCmd, LPNMTOOLBAR pnmtb);
	void    SyncFoldersCheckState();  // called by BagWriteHook / ToolbarSubclassProc
	bool    EnsureBrowserBag();       // lazy-init bag + hook; returns true if bag was just acquired
	CAddressBar& GetAddressBar() { return m_addressBar; }
	void    LayoutBandChildren();

	DECLARE_REGISTRY_RESOURCEID_V2_WITHOUT_MODULE(IDR_CLASSICEXPLORER, CStandardToolbar)

	BEGIN_SINK_MAP(CStandardToolbar)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_ONQUIT, OnQuit)
		SINK_ENTRY_EX(1, DIID_DWebBrowserEvents2, DISPID_COMMANDSTATECHANGE, OnCommandStateChange)
	END_SINK_MAP()

	BEGIN_COM_MAP(CStandardToolbar)
		COM_INTERFACE_ENTRY(IOleWindow)
		COM_INTERFACE_ENTRY(IObjectWithSite)
		COM_INTERFACE_ENTRY_IID(IID_IDockingWindow, IDockingWindow)
		COM_INTERFACE_ENTRY_IID(IID_IDeskBand, IDeskBand)
		COM_INTERFACE_ENTRY_IID(IID_IDeskBand2, IDeskBand2)
		COM_INTERFACE_ENTRY_IID(IID_IInputObject, IInputObject)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct() { return S_OK; }
	void FinalRelease() {}

	// IDeskBand:
	STDMETHOD(GetBandInfo)(DWORD dwBandId, DWORD dwViewMode, DESKBANDINFO* pDbi);

	// IDeskBand2:
	STDMETHOD(CanRenderComposited)(BOOL* pfCanRenderComposited);
	STDMETHOD(SetCompositionState)(BOOL fCompositionEnabled);
	STDMETHOD(GetCompositionState)(BOOL* pfCompositionEnabled);

	// IObjectWithSite:
	STDMETHOD(SetSite)(IUnknown* pUnkSite);

	// IOleWindow:
	STDMETHOD(GetWindow)(HWND* hWnd);
	STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode);

	// IDockingWindow:
	STDMETHOD(CloseDW)(unsigned long dwReserved);
	STDMETHOD(ResizeBorderDW)(const RECT* pRcBorder, IUnknown* pUnkToolbarSite, BOOL fReserved);
	STDMETHOD(ShowDW)(BOOL fShow);

	// DWebBrowserEvents2:
	STDMETHOD(OnNavigateComplete)(IDispatch* pDisp, VARIANT* url);
	STDMETHOD(OnQuit)();
	STDMETHOD_(void, OnCommandStateChange)(long Command, VARIANT_BOOL Enable);

	// IInputObject:
	STDMETHOD(HasFocusIO)();
	STDMETHOD(TranslateAcceleratorIO)(MSG* pMsg);
	STDMETHOD(UIActivateIO)(BOOL fActivate, MSG* pMsg);
};

OBJECT_ENTRY_AUTO(__uuidof(CStandardToolbar), CStandardToolbar);

#endif // _STANDARDTOOLBAR_H
