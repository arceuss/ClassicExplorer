/*
 * StandardToolbar.cpp: Implements the Windows XP-style Standard Buttons toolbar band.
 *
 * Ported from the Windows Server 2003 source code (shell\browseui\itbar.cpp and
 * shell\shell32\defview.cpp). The toolbar is registered as an Explorer bar (IDeskBand2)
 * and appears in the rebar alongside the address bar and links bar.
 *
 * ================================================================================================
 * ---- IMPORTANT FUNCTIONS ----
 *
 *  - SetSite:              Installs / removes the toolbar.
 *  - InitToolbarButtons:   Populates the toolbar with buttons.
 *  - UpdateButtonStates:   Enables/disables buttons based on current state.
 *  - OnNavigateComplete:   Called after each navigation to refresh button states.
 */

#include "stdafx.h"
#include "framework.h"
#include "resource.h"
#include "ClassicExplorer_i.h"
#include "dllmain.h"

#include "StandardToolbar.h"
#include "../util/util.h"

#include <commctrl.h>
#include <commoncontrols.h>
#include <shlwapi.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "mpr.lib")

// OLECMDID values for back/forward navigation (may not be in all SDK versions)
#ifndef OLECMDID_NAVIGATEBACKWARD
#define OLECMDID_NAVIGATEBACKWARD 63
#endif
#ifndef OLECMDID_NAVIGATEFORWARD
#define OLECMDID_NAVIGATEFORWARD 64
#endif

// ================================================================================================
// Travel Log interfaces (from IE Platform SDK tlogstg.h)
// These are not included in the modern Windows SDK but are still supported by Explorer.
//

#ifndef __ITravelLogEntry_INTERFACE_DEFINED__
#define __ITravelLogEntry_INTERFACE_DEFINED__
MIDL_INTERFACE("7EBFDD87-AD18-11d3-A4C5-00C04F72D6B8")
ITravelLogEntry : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE GetTitle(LPOLESTR *ppszTitle) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetURL(LPOLESTR *ppszURL) = 0;
};
#endif

#ifndef __IEnumTravelLogEntry_INTERFACE_DEFINED__
#define __IEnumTravelLogEntry_INTERFACE_DEFINED__
MIDL_INTERFACE("7EBFDD85-AD18-11d3-A4C5-00C04F72D6B8")
IEnumTravelLogEntry : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE Next(ULONG cElt, ITravelLogEntry **rgElt, ULONG *pcEltFetched) = 0;
    virtual HRESULT STDMETHODCALLTYPE Skip(ULONG cElt) = 0;
    virtual HRESULT STDMETHODCALLTYPE Reset() = 0;
    virtual HRESULT STDMETHODCALLTYPE Clone(IEnumTravelLogEntry **ppEnum) = 0;
};
#endif

#ifndef __ITravelLogStg_INTERFACE_DEFINED__
#define __ITravelLogStg_INTERFACE_DEFINED__

#define TLEF_RELATIVE_BACK    0x00000010
#define TLEF_RELATIVE_FORE    0x00000020

MIDL_INTERFACE("7EBFDD80-AD18-11d3-A4C5-00C04F72D6B8")
ITravelLogStg : public IUnknown
{
public:
    virtual HRESULT STDMETHODCALLTYPE CreateEntry(LPCOLESTR pszUrl, LPCOLESTR pszTitle,
        ITravelLogEntry *ptleRelativeTo, BOOL fPrepend, ITravelLogEntry **pptle) = 0;
    virtual HRESULT STDMETHODCALLTYPE TravelTo(ITravelLogEntry *ptle) = 0;
    virtual HRESULT STDMETHODCALLTYPE EnumEntries(DWORD flags, IEnumTravelLogEntry **ppenum) = 0;
    virtual HRESULT STDMETHODCALLTYPE FindEntries(DWORD flags, LPCOLESTR pszUrl,
        IEnumTravelLogEntry **ppenum) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCount(DWORD flags, DWORD *pcEntries) = 0;
    virtual HRESULT STDMETHODCALLTYPE RemoveEntry(ITravelLogEntry *ptle) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetRelativeEntry(int iOffset, ITravelLogEntry **ptle) = 0;
};

#define IID_ITravelLogStg    __uuidof(ITravelLogStg)
#define SID_STravelLogCursor __uuidof(ITravelLogStg)
#endif

// Service IDs used to obtain the per-browser property bag (from Open-Shell)
static const GUID SID_FrameManager =
	{ 0x31e4fa78, 0x02b4, 0x419f, { 0x94, 0x30, 0x7b, 0x75, 0x85, 0x23, 0x7c, 0x77 } };
static const GUID SID_PerBrowserPropertyBag =
	{ 0xa3b24a0a, 0x7b68, 0x448d, { 0x99, 0x79, 0xc7, 0x00, 0x05, 0x9c, 0x3a, 0xd1 } };

// Property bag key for the navigation pane visibility
static const WCHAR g_NavPaneVisible[] = L"PageSpaceControlSizer_Visible";

// Private message posted to the toolbar when the property bag Write hook fires.
#define WM_BAGWRITE_NOTIFY  (WM_APP + 1)

// ================================================================================================
// IPropertyBag::Write vtable hook (same technique as Open-Shell ExplorerBand)
//
// We patch the vtable of the per-browser property bag so that every Write() call
// is intercepted.  When anything (DUI sizer, Windhawk mod, etc.) writes
// PageSpaceControlSizer_Visible, we post a message to our toolbar so it can
// re-read the value and sync the Folders check-button — no timers, no guessing.
//

typedef HRESULT (__stdcall *tBagWrite)(IPropertyBag* pThis, LPCOLESTR pszPropName, VARIANT* pVar);
static volatile tBagWrite g_OldBagWrite = nullptr;
static volatile HWND     g_hWndToolbarForHook = nullptr;  // toolbar hwnd that receives WM_BAGWRITE_NOTIFY

static HRESULT __stdcall BagWriteHook(IPropertyBag* pThis, LPCOLESTR pszPropName, VARIANT* pVar)
{
	HRESULT hr = g_OldBagWrite(pThis, pszPropName, pVar);
	if (_wcsicmp(pszPropName, g_NavPaneVisible) == 0)
	{
		HWND hWnd = g_hWndToolbarForHook;
		if (hWnd && IsWindow(hWnd))
			PostMessage(hWnd, WM_BAGWRITE_NOTIFY, 0, 0);
	}
	return hr;
}

// WM_COMMAND IDs for SHELLDLL_DefView (from Open-Shell, verified on Win7-11)
#define DEFVIEW_COPYTO      28702
#define DEFVIEW_MOVETO      28703
#define DEFVIEW_DELETE       28689
#define DEFVIEW_PROPERTIES  28691   // SFVIDM_FILE_PROPERTIES
#define DEFVIEW_CUT         28696   // SFVIDM_EDIT_CUT
#define DEFVIEW_COPY        28697   // SFVIDM_EDIT_COPY
#define DEFVIEW_PASTE       28698   // SFVIDM_EDIT_PASTE

// WM_COMMAND IDs sent to ShellTabWindowClass
#define SHELLTAB_UNDO           28699
#define SHELLTAB_REFRESH        41504
#define SHELLTAB_MAPDRIVE       41089
#define SHELLTAB_DISCONNECT     41090
#define SHELLTAB_FOLDEROPTIONS  28771   // SFVIDM_TOOL_OPTIONS

// View mode WM_COMMAND IDs for SHELLDLL_DefView (from shell32 MUI resources)
// XP+ view modes (from SFVIDM_VIEW_FIRST + 0x0F range)
#define DEFVIEW_VIEW_THUMBNAILS  28751
#define DEFVIEW_VIEW_TILES       28748
#define DEFVIEW_VIEW_ICONS       28750
#define DEFVIEW_VIEW_SMALLICONS  28752
#define DEFVIEW_VIEW_LIST        28753
#define DEFVIEW_VIEW_DETAILS     28747

// ================================================================================================
// Registered window message for settings change notification (broadcast by Customize dialog)
//
static UINT g_uMsgSettingsChanged = 0;

static UINT GetSettingsChangedMsg()
{
	if (!g_uMsgSettingsChanged)
		g_uMsgSettingsChanged = RegisterWindowMessageW(CE_WM_SETTINGS_CHANGED_NAME);
	return g_uMsgSettingsChanged;
}

// ================================================================================================
// Internal comctl32 structure for TBN_INITCUSTOMIZE / TBN_RESET notifications.
// Not in public SDK headers but present in comctl32 v5/v6 since Windows XP.
//

typedef struct tagNMTBCUSTOMIZEDLG {
	NMHDR hdr;
	HWND  hDlg;
} NMTBCUSTOMIZEDLG, *LPNMTBCUSTOMIZEDLG;

#ifndef TBNRF_HIDEHELP
#define TBNRF_HIDEHELP      0x00000001
#endif

#ifndef TBNRF_ENDCUSTOMIZE
#define TBNRF_ENDCUSTOMIZE  0x00000002
#endif

// ================================================================================================
// Registered window message for customize toolbar request (sent by BrandBand menu)
//

static UINT g_uMsgCustomizeToolbar = 0;

static UINT GetCustomizeToolbarMsg()
{
	if (!g_uMsgCustomizeToolbar)
		g_uMsgCustomizeToolbar = RegisterWindowMessageW(CE_WM_CUSTOMIZE_TOOLBAR_NAME);
	return g_uMsgCustomizeToolbar;
}

// ================================================================================================
// Customize Toolbar child dialog (text/icon options)
// Embedded in comctl32's TB_CUSTOMIZE dialog via TBN_INITCUSTOMIZE.
// Ported from Windows XP's DLG_TEXTICONOPTIONS / _BtnAttrDlgProc in itbar.cpp.
//

static void _PopulateCombo(HWND hwndCombo, const int* pIds, int cIds, HINSTANCE hInst)
{
	for (int i = 0; i < cIds; i++)
	{
		WCHAR szText[64] = {};
		LoadStringW(hInst, pIds[i], szText, ARRAYSIZE(szText));
		int idx = (int)SendMessageW(hwndCombo, CB_ADDSTRING, 0, (LPARAM)szText);
		SendMessageW(hwndCombo, CB_SETITEMDATA, idx, (LPARAM)pIds[i]);
	}
}

static void _SetComboSelection(HWND hwndCombo, int idsTarget)
{
	int cItems = (int)SendMessageW(hwndCombo, CB_GETCOUNT, 0, 0);
	for (int i = 0; i < cItems; i++)
	{
		if ((int)SendMessageW(hwndCombo, CB_GETITEMDATA, i, 0) == idsTarget)
		{
			SendMessageW(hwndCombo, CB_SETCURSEL, i, 0);
			return;
		}
	}
	SendMessageW(hwndCombo, CB_SETCURSEL, 0, 0);
}

static const int c_iTextOptions[] = { IDS_TEXTLABELS, IDS_PARTIALTEXT, IDS_NOTEXTLABELS };
static const int c_iIconOptions[] = { IDS_LARGEICONS, IDS_SMALLICONS };

static DWORD _TextIdsToMode(INT_PTR ids)
{
	switch (ids)
	{
	case IDS_TEXTLABELS:  return CE_TEXTMODE_SHOWTEXT;
	case IDS_PARTIALTEXT: return CE_TEXTMODE_SELECTIVE;
	case IDS_NOTEXTLABELS:return CE_TEXTMODE_NOTEXT;
	default:              return CE_TEXTMODE_SELECTIVE;
	}
}

static int _TextModeToIds(DWORD dwMode)
{
	switch (dwMode)
	{
	case CE_TEXTMODE_SHOWTEXT:  return IDS_TEXTLABELS;
	case CE_TEXTMODE_SELECTIVE: return IDS_PARTIALTEXT;
	case CE_TEXTMODE_NOTEXT:    return IDS_NOTEXTLABELS;
	default:                    return IDS_PARTIALTEXT;
	}
}

static INT_PTR CALLBACK CustomizeChildDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
	{
		HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
		CEUtil::CESettings settings = CEUtil::GetCESettings();

		HWND hwndText = GetDlgItem(hDlg, IDC_SHOWTEXT);
		HWND hwndIcons = GetDlgItem(hDlg, IDC_SMALLICONS);

		_PopulateCombo(hwndText, c_iTextOptions, ARRAYSIZE(c_iTextOptions), hInst);
		_PopulateCombo(hwndIcons, c_iIconOptions, ARRAYSIZE(c_iIconOptions), hInst);

		_SetComboSelection(hwndText, _TextModeToIds(settings.textLabelMode));
		_SetComboSelection(hwndIcons, settings.smallIcons ? IDS_SMALLICONS : IDS_LARGEICONS);

		// IE 5.5 style and Win98 Views checkboxes — only visible when 2K skin is active
		HWND hwndIE55 = GetDlgItem(hDlg, IDC_IE55STYLE);
		HWND hwndW98V = GetDlgItem(hDlg, IDC_WIN98VIEWS);
		if (settings.theme == CLASSIC_EXPLORER_2K)
		{
			ShowWindow(hwndIE55, SW_SHOW);
			SendMessage(hwndIE55, BM_SETCHECK, settings.ie55Style ? BST_CHECKED : BST_UNCHECKED, 0);
			ShowWindow(hwndW98V, SW_SHOW);
			SendMessage(hwndW98V, BM_SETCHECK, settings.win98Views == 1 ? BST_CHECKED : BST_UNCHECKED, 0);
		}
		else
		{
			ShowWindow(hwndIE55, SW_HIDE);
			ShowWindow(hwndW98V, SW_HIDE);
		}

		return TRUE;
	}

	case WM_COMMAND:
	{
		int idCtrl = LOWORD(wParam);
		int notif = HIWORD(wParam);

		if ((idCtrl == IDC_SHOWTEXT || idCtrl == IDC_SMALLICONS) &&
			(notif == CBN_SELENDOK || notif == CBN_CLOSEUP))
		{
			HWND hwndText = GetDlgItem(hDlg, IDC_SHOWTEXT);
			HWND hwndIcons = GetDlgItem(hDlg, IDC_SMALLICONS);

			INT_PTR iTextSel = SendMessageW(hwndText, CB_GETCURSEL, 0, 0);
			INT_PTR idsText = SendMessageW(hwndText, CB_GETITEMDATA, iTextSel, 0);
			DWORD dwTextMode = _TextIdsToMode(idsText);

			INT_PTR iIconSel = SendMessageW(hwndIcons, CB_GETCURSEL, 0, 0);
			INT_PTR idsIcon = SendMessageW(hwndIcons, CB_GETITEMDATA, iIconSel, 0);
			DWORD dwSmallIcons = (idsIcon == IDS_SMALLICONS) ? 1 : 0;

			// Save to registry — changes take effect in newly opened Explorer windows
			CEUtil::CESettings newSettings(CLASSIC_EXPLORER_NONE, -1, -1, -1, dwSmallIcons, dwTextMode);
			CEUtil::WriteCESettings(newSettings);

			return TRUE;
		}

		if (idCtrl == IDC_IE55STYLE && notif == BN_CLICKED)
		{
			DWORD dwIE55 = (SendMessage(GetDlgItem(hDlg, IDC_IE55STYLE), BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
			CEUtil::CESettings newSettings(CLASSIC_EXPLORER_NONE, -1, -1, -1, -1, -1, dwIE55);
			CEUtil::WriteCESettings(newSettings);
			return TRUE;
		}

		if (idCtrl == IDC_WIN98VIEWS && notif == BN_CLICKED)
		{
			DWORD dwW98V = (SendMessage(GetDlgItem(hDlg, IDC_WIN98VIEWS), BM_GETCHECK, 0, 0) == BST_CHECKED) ? 1 : 0;
			CEUtil::CESettings newSettings(CLASSIC_EXPLORER_NONE, -1, -1, -1, -1, -1, -1, dwW98V);
			CEUtil::WriteCESettings(newSettings);
			return TRUE;
		}
		break;
	}
	}
	return FALSE;
}

// ================================================================================================
// Top-level window subclass to receive broadcast settings change messages
// and forward them to the CStandardToolbar instance.
//

static LRESULT CALLBACK TopLevelSubclassProc(
	HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
	UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	CStandardToolbar* pThis = reinterpret_cast<CStandardToolbar*>(dwRefData);

	if (uMsg == GetSettingsChangedMsg())
	{
		pThis->ReloadSettings();
		return 0;
	}

	if (uMsg == GetCustomizeToolbarMsg())
	{
		SendMessage(pThis->GetToolbarHwnd(), TB_CUSTOMIZE, 0, 0);
		return 0;
	}

	if (uMsg == WM_NCDESTROY)
	{
		RemoveWindowSubclass(hWnd, TopLevelSubclassProc, uIdSubclass);
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// ================================================================================================
// Subclass procedure for the toolbar to relay messages to parent band
//

static LRESULT CALLBACK ToolbarSubclassProc(
	HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
	UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	CStandardToolbar* pThis = reinterpret_cast<CStandardToolbar*>(dwRefData);

	switch (uMsg)
	{
	case WM_PAINT:
		// Lazy-init the property bag on first paint (same pattern as Open-Shell).
		// By WM_PAINT time the DUI has written PageSpaceControlSizer_Visible,
		// so getting the bag now installs the vtable hook with a valid HWND
		// and the subsequent sync reads the correct initial value.
		if (pThis->EnsureBrowserBag())
			PostMessage(hWnd, WM_BAGWRITE_NOTIFY, 0, 0);
		break;

	case WM_BAGWRITE_NOTIFY:
		// The property bag Write hook fired for PageSpaceControlSizer_Visible.
		// Re-read the actual value and sync the Folders button.
		pThis->SyncFoldersCheckState();
		return 0;

	case WM_NCDESTROY:
		g_hWndToolbarForHook = nullptr;
		RemoveWindowSubclass(hWnd, ToolbarSubclassProc, uIdSubclass);
		break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// ================================================================================================
// IDeskBand2 implementation
//

STDMETHODIMP CStandardToolbar::CanRenderComposited(BOOL* pfCanRenderComposited)
{
	if (pfCanRenderComposited)
		*pfCanRenderComposited = TRUE;
	return S_OK;
}

STDMETHODIMP CStandardToolbar::SetCompositionState(BOOL fCompositionEnabled)
{
	return S_OK;
}

STDMETHODIMP CStandardToolbar::GetCompositionState(BOOL* pfCompositionEnabled)
{
	if (pfCompositionEnabled)
		*pfCompositionEnabled = TRUE;
	return S_OK;
}

// ================================================================================================
// IDeskBand implementation
//

STDMETHODIMP CStandardToolbar::GetBandInfo(DWORD dwBandId, DWORD dwViewMode, DESKBANDINFO* pDbi)
{
	if (!pDbi)
		return E_INVALIDARG;

	// Match XP's TOOLSBANDCLASS::GetBandInfo:
	// - ptMinSize.x = width of first button (minimum to show something useful)
	// - ptMinSize.y = height of first button
	// - dwModeFlags includes DBIMF_USECHEVRON for overflow handling

	pDbi->dwModeFlags = DBIMF_USECHEVRON;

	if (pDbi->dwMask & DBIM_MINSIZE)
	{
		if (m_hWndToolbar && SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0))
		{
			RECT rc = {};
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, 0, (LPARAM)&rc);
			pDbi->ptMinSize.x = rc.right - rc.left;
			pDbi->ptMinSize.y = rc.bottom - rc.top;
		}
		else
		{
			LONG lSize = (LONG)SendMessage(m_hWndToolbar, TB_GETBUTTONSIZE, 0, 0);
			pDbi->ptMinSize.x = LOWORD(lSize);
			pDbi->ptMinSize.y = HIWORD(lSize);
		}
	}

	if (pDbi->dwMask & DBIM_MAXSIZE)
	{
		pDbi->ptMaxSize.x = 0;
		pDbi->ptMaxSize.y = -1;
	}

	if (pDbi->dwMask & DBIM_INTEGRAL)
	{
		pDbi->ptIntegral.x = 0;
		pDbi->ptIntegral.y = 1;
	}

	if (pDbi->dwMask & DBIM_ACTUAL)
	{
		SIZE szIdeal = {};
		if (m_hWndToolbar)
		{
			SendMessage(m_hWndToolbar, TB_GETIDEALSIZE, FALSE, (LPARAM)&szIdeal);
			RECT rc = {};
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, 0, (LPARAM)&rc);
			pDbi->ptActual.x = szIdeal.cx;
			pDbi->ptActual.y = rc.bottom - rc.top;
		}
	}

	if (pDbi->dwMask & DBIM_TITLE)
	{
		pDbi->dwMask &= ~DBIM_TITLE;
	}

	if (pDbi->dwMask & DBIM_BKCOLOR)
	{
		pDbi->dwMask &= ~DBIM_BKCOLOR;
	}

	return S_OK;
}

// ================================================================================================
// IOleWindow implementation
//

STDMETHODIMP CStandardToolbar::GetWindow(HWND* hWnd)
{
	if (!hWnd)
		return E_INVALIDARG;

	*hWnd = m_hWndToolbar;
	return S_OK;
}

STDMETHODIMP CStandardToolbar::ContextSensitiveHelp(BOOL fEnterMode)
{
	return S_OK;
}

// ================================================================================================
// IDockingWindow implementation
//

STDMETHODIMP CStandardToolbar::CloseDW(unsigned long dwReserved)
{
	return ShowDW(FALSE);
}

STDMETHODIMP CStandardToolbar::ResizeBorderDW(const RECT* pRcBorder, IUnknown* pUnkToolbarSite, BOOL fReserved)
{
	return E_NOTIMPL;
}

STDMETHODIMP CStandardToolbar::ShowDW(BOOL fShow)
{
	m_bShow = (fShow != FALSE);
	if (m_hWndToolbar)
		ShowWindow(m_hWndToolbar, fShow ? SW_SHOW : SW_HIDE);
	return S_OK;
}

// ================================================================================================
// IObjectWithSite implementation
//

STDMETHODIMP CStandardToolbar::SetSite(IUnknown* pUnkSite)
{
	// Let ATL store the site
	IObjectWithSiteImpl<CStandardToolbar>::SetSite(pUnkSite);

	if (pUnkSite)
	{
		// Obtain the parent window
		CComQIPtr<IOleWindow> pOleWindow = pUnkSite;
		if (pOleWindow)
			pOleWindow->GetWindow(&m_hWndParent);

		if (!IsWindow(m_hWndParent))
			return E_FAIL;

		// Obtain the input object site for focus negotiation
		pUnkSite->QueryInterface(IID_IInputObjectSite, (void**)&m_pSite);

		// Obtain shell browser and web browser interfaces via IServiceProvider
		CComQIPtr<IServiceProvider> pProvider = pUnkSite;
		if (pProvider)
		{
			pProvider->QueryService(SID_SShellBrowser, IID_IShellBrowser, (void**)&m_pShellBrowser);
			pProvider->QueryService(SID_SWebBrowserApp, IID_IWebBrowser2, (void**)&m_pWebBrowser);
		}

		// Create the toolbar
		HRESULT hr = CreateToolbarWindow(m_hWndParent);
		if (FAILED(hr))
			return hr;

		// Advise for browser events
		if (m_pWebBrowser)
		{
			if (m_dwEventCookie == 0xFEFEFEFE)
				DispEventAdvise(m_pWebBrowser, &DIID_DWebBrowserEvents2);
		}

		// Initial button state update
		UpdateButtonStates();

		// Subclass the top-level Explorer window to receive settings change broadcasts
		HWND hTop = GetAncestor(m_hWndParent, GA_ROOT);
		if (hTop)
			SetWindowSubclass(hTop, TopLevelSubclassProc, (UINT_PTR)this, reinterpret_cast<DWORD_PTR>(this));
	}
	else
	{
		// Site is being removed - clean up
		// Remove top-level subclass
		if (m_hWndParent)
		{
			HWND hTop = GetAncestor(m_hWndParent, GA_ROOT);
			if (hTop)
				RemoveWindowSubclass(hTop, TopLevelSubclassProc, (UINT_PTR)this);
		}

		if (m_pWebBrowser && m_dwEventCookie != 0xFEFEFEFE)
			DispEventUnadvise(m_pWebBrowser, &DIID_DWebBrowserEvents2);

		DestroyToolbarWindow();

		g_hWndToolbarForHook = nullptr;
		m_pSite = nullptr;
		m_pWebBrowser.Release();
		m_pShellBrowser.Release();
		m_pBrowserBag.Release();
		m_hWndParent = nullptr;
	}

	return S_OK;
}

// ================================================================================================
// DWebBrowserEvents2 handlers
//

STDMETHODIMP CStandardToolbar::OnNavigateComplete(IDispatch* pDisp, VARIANT* url)
{
	UpdateButtonStates();
	return S_OK;
}

STDMETHODIMP CStandardToolbar::OnQuit()
{
	if (m_pWebBrowser && m_dwEventCookie != 0xFEFEFEFE)
		return DispEventUnadvise(m_pWebBrowser, &DIID_DWebBrowserEvents2);
	return S_OK;
}

void __stdcall CStandardToolbar::OnCommandStateChange(long Command, VARIANT_BOOL Enable)
{
	// CSC_NAVIGATEBACK = 2, CSC_NAVIGATEFORWARD = 1
	if (Command == 2)
	{
		m_bCanGoBack = (Enable != VARIANT_FALSE);
		if (m_hWndToolbar)
			SendMessage(m_hWndToolbar, TB_ENABLEBUTTON, TBIDM_BACK, MAKELONG(m_bCanGoBack, 0));
	}
	else if (Command == 1)
	{
		m_bCanGoForward = (Enable != VARIANT_FALSE);
		if (m_hWndToolbar)
			SendMessage(m_hWndToolbar, TB_ENABLEBUTTON, TBIDM_FORWARD, MAKELONG(m_bCanGoForward, 0));
	}
}

// ================================================================================================
// IInputObject implementation
//

STDMETHODIMP CStandardToolbar::HasFocusIO()
{
	HWND hFocus = GetFocus();
	if (hFocus == m_hWndToolbar || IsChild(m_hWndToolbar, hFocus))
		return S_OK;
	return S_FALSE;
}

STDMETHODIMP CStandardToolbar::TranslateAcceleratorIO(MSG* pMsg)
{
	return S_FALSE;
}

STDMETHODIMP CStandardToolbar::UIActivateIO(BOOL fActivate, MSG* pMsg)
{
	if (fActivate && m_hWndToolbar)
		SetFocus(m_hWndToolbar);
	return S_OK;
}

// ================================================================================================
// Toolbar creation and initialization
//

// Helper: Load a bitmap resource and create an image list matching the XP approach.
// Uses LR_CREATEDIBSECTION so that 32bpp alpha bitmaps are loaded as DIB sections,
// and the actual bit depth is detected from the bitmap to set ILC flags correctly.
// This is a port of browseui's CreateImageList() from util.cpp.

// Private comctl32 flag used to extract color depth from bmBitsPixel.
#ifndef ILC_COLORMASK
#define ILC_COLORMASK 0x00FE
#endif

static HIMAGELIST _CreateImageListFromResource(HINSTANCE hInst, UINT idBitmap, int cx, COLORREF crMask)
{
	HBITMAP hbm = (HBITMAP)LoadImage(hInst, MAKEINTRESOURCE(idBitmap),
		IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
	if (!hbm)
		return nullptr;

	BITMAP bm = {};
	if (GetObject(hbm, sizeof(bm), &bm) != sizeof(bm))
	{
		DeleteObject(hbm);
		return nullptr;
	}

	int cy = bm.bmHeight;
	int cInitial = bm.bmWidth / cx;

	// Determine ILC flags from the bitmap's actual bit depth
	UINT flags = ILC_MASK;
	if (bm.bmBits)
		flags |= (bm.bmBitsPixel & ILC_COLORMASK);
	else
		flags |= ILC_COLOR24;

	HIMAGELIST himl = ImageList_Create(cx, cy, flags, cInitial, 0);
	if (himl)
	{
		if (ImageList_AddMasked(himl, hbm, crMask) < 0)
		{
			ImageList_Destroy(himl);
			himl = nullptr;
		}
	}

	DeleteObject(hbm);
	return himl;
}

// ================================================================================================
// XP-style grayscale image list for disabled button state.
// Ported from Windows Server 2003 browseui itbar.cpp _CreateGrayScaleImagelist().
//
// Algorithm matches comctl32 v6's ILS_SATURATE + AlphaBlend path used when
// the toolbar has no explicit disabled image list:
//   1. TrueSaturateBits: gray = (54*R + 183*G + 19*B) >> 8  (weighted luminance)
//   2. AlphaBlend with SourceConstantAlpha = 150 (59% opacity)
//
// We pre-apply both steps into the bitmap so the toolbar can use it directly.

static HIMAGELIST _CreateDisabledImageList(HINSTANCE hInst, UINT idBitmap, int cx, COLORREF crMask,
	ClassicExplorerTheme theme = CLASSIC_EXPLORER_XP)
{
	// Load a fresh copy of the bitmap as a DIB section (preserves native format)
	HBITMAP hbm = (HBITMAP)LoadImage(hInst, MAKEINTRESOURCE(idBitmap),
		IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
	if (!hbm)
		return nullptr;

	BITMAP bm = {};
	if (GetObject(hbm, sizeof(bm), &bm) != sizeof(bm) || !bm.bmBits)
	{
		DeleteObject(hbm);
		return nullptr;
	}

	int totalWidth  = bm.bmWidth;
	int totalHeight = bm.bmHeight;
	int cImages     = totalWidth / cx;
	int bpp         = bm.bmBitsPixel;

	if (theme == CLASSIC_EXPLORER_2K)
	{
		// Win2K browseui never creates a disabled image list — it relies on comctl32's
		// built-in PSDPxax emboss (etched shadow+highlight) rendered at draw time.
		// Modern comctl32 v6 still has this code path for non-alpha (8bpp/24bpp) images,
		// so we just return nullptr to let the toolbar handle disabled rendering natively.
		DeleteObject(hbm);
		return nullptr;
	}
	else
	{
		// XP disabled algorithm: TrueSaturateBits weighted luminance + 59% alpha
		if (bpp == 32)
		{
			BYTE* pBits = (BYTE*)bm.bmBits;
			long cbScan = bm.bmWidthBytes;
			for (int y = 0; y < totalHeight; y++)
			{
				DWORD* pPixel = (DWORD*)(pBits + y * cbScan);
				for (int x = 0; x < totalWidth; x++, pPixel++)
				{
					DWORD px = *pPixel;
					BYTE a = (BYTE)(px >> 24);
					if (a == 0) continue;

					BYTE b = (BYTE)(px);
					BYTE g = (BYTE)(px >> 8);
					BYTE r = (BYTE)(px >> 16);

					// comctl32 v6 TrueSaturateBits weighted luminance
					BYTE gray = (BYTE)((54u * r + 183u * g + 19u * b) >> 8);

					// Pre-multiply with SourceConstantAlpha = 150 (59% opacity)
					BYTE newA = (BYTE)((a * 150u) / 255u);

					*pPixel = ((DWORD)newA << 24) | ((DWORD)gray << 16) | ((DWORD)gray << 8) | gray;
				}
			}
		}
		else if (bpp == 24)
		{
			BYTE* pBits = (BYTE*)bm.bmBits;
			long cbScan = bm.bmWidthBytes;
			BYTE maskR = GetRValue(crMask), maskG = GetGValue(crMask), maskB = GetBValue(crMask);
			for (int y = 0; y < totalHeight; y++)
			{
				BYTE* pScan = pBits + y * cbScan;
				for (int x = 0; x < totalWidth; x++, pScan += 3)
				{
					if (pScan[0] == maskB && pScan[1] == maskG && pScan[2] == maskR)
						continue;
					BYTE gray = (BYTE)((54u * pScan[2] + 183u * pScan[1] + 19u * pScan[0]) >> 8);
					pScan[0] = pScan[1] = pScan[2] = gray;
				}
			}
		}
	}

	// Create image list matching the bitmap's native bit depth
	UINT flags = ILC_MASK | (bpp & ILC_COLORMASK);
	HIMAGELIST himl = ImageList_Create(cx, totalHeight, flags, cImages, 0);
	if (himl)
	{
		if (ImageList_AddMasked(himl, hbm, crMask) < 0)
		{
			ImageList_Destroy(himl);
			himl = nullptr;
		}
	}

	DeleteObject(hbm);
	return himl;
}

// ================================================================================================
// Append icons from shell toolbar bitmaps (XP or 2K) to our image lists.
// The shell32 bitmaps (47 glyphs, browseui shdef layout) contain icons for Properties,
// Cut, Copy, Paste, and Folder Options at known positions. We extract just those 5
// and append them.
//
// Shell32 bitmap glyph layout (shared between XP and 2K):
//   Index 5  = STD_CUT       (OFFSET_STD + STD_CUT)
//   Index 6  = STD_COPY      (OFFSET_STD + STD_COPY)
//   Index 7  = STD_PASTE     (OFFSET_STD + STD_PASTE)
//   Index 15 = STD_PROPERTIES(OFFSET_STD + STD_PROPERTIES)
//   Index 46 = VIEW_OPTIONS  (OFFSET_VIEW + VIEW_OPTIONS)
//

static void _AppendShell32Icons(HIMAGELIST himlTarget, UINT idShell32Bmp, int cx, COLORREF crMask,
	const int* pSrcIndices, int cIndices)
{
	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	HBITMAP hbm = (HBITMAP)LoadImage(hInst, MAKEINTRESOURCE(idShell32Bmp),
		IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION);
	if (!hbm) return;

	BITMAP bm = {};
	GetObject(hbm, sizeof(bm), &bm);
	int cy = bm.bmHeight;

	// Create a temporary image list from the full shell32 bitmap
	UINT flags = ILC_MASK | (bm.bmBitsPixel & ILC_COLORMASK);
	HIMAGELIST himlSrc = ImageList_Create(cx, cy, flags, bm.bmWidth / cx, 0);
	if (himlSrc)
	{
		ImageList_AddMasked(himlSrc, hbm, crMask);

		for (int i = 0; i < cIndices; i++)
		{
			HICON hIcon = ImageList_GetIcon(himlSrc, pSrcIndices[i], ILD_TRANSPARENT);
			if (hIcon)
			{
				ImageList_AddIcon(himlTarget, hIcon);
				DestroyIcon(hIcon);
			}
		}
		ImageList_Destroy(himlSrc);
	}
	DeleteObject(hbm);
}

// Shell32 glyph source indices for our 5 extra buttons
// Order must match GLYPHIDX_PROPERTIES(18), CUT(19), COPY(20), PASTE(21), FOLDEROPTIONS(22)
static const int c_shell32GlyphIndices[] = { 15, 5, 6, 7, 46 };
static const int c_nShell32Glyphs = ARRAYSIZE(c_shell32GlyphIndices);

static void _AppendSystemBitmapIcons(HIMAGELIST himlDef, HIMAGELIST himlHot, HIMAGELIST himlDis, int cx, ClassicExplorerTheme theme, bool fIE55 = false)
{
	COLORREF crMask = RGB(255, 0, 255);
	UINT idDef, idHot;
	if (theme == CLASSIC_EXPLORER_2K)
	{
		if (fIE55)
		{
			if (cx <= 16) { idDef = IDB_2K_IE55_SHELL32_DEF_16; idHot = IDB_2K_IE55_SHELL32_HOT_16; }
			else          { idDef = IDB_2K_IE55_SHELL32_DEF_20; idHot = IDB_2K_IE55_SHELL32_HOT_20; }
		}
		else
		{
			if (cx <= 16) { idDef = IDB_2K_SHELL32_DEF_16; idHot = IDB_2K_SHELL32_HOT_16; }
			else          { idDef = IDB_2K_SHELL32_DEF_20; idHot = IDB_2K_SHELL32_HOT_20; }
		}
	}
	else
	{
		if (cx <= 16) { idDef = IDB_SHELL32_DEF_16; idHot = IDB_SHELL32_HOT_16; }
		else          { idDef = IDB_SHELL32_DEF_24; idHot = IDB_SHELL32_HOT_24; }
	}

	if (himlDef) _AppendShell32Icons(himlDef, idDef, cx, crMask, c_shell32GlyphIndices, c_nShell32Glyphs);
	if (himlHot) _AppendShell32Icons(himlHot, idHot, cx, crMask, c_shell32GlyphIndices, c_nShell32Glyphs);
	if (himlDis) _AppendShell32Icons(himlDis, idDef, cx, crMask, c_shell32GlyphIndices, c_nShell32Glyphs);
}

HRESULT CStandardToolbar::CreateToolbarWindow(HWND hWndParent)
{
	// Read current settings (icon size, text mode, theme)
	CEUtil::CESettings settings = CEUtil::GetCESettings();
	bool fSmallIcons = (settings.smallIcons == 1);
	DWORD dwTextMode = (settings.textLabelMode <= 2) ? settings.textLabelMode : CE_TEXTMODE_SELECTIVE;

	// Determine initial window style based on text mode:
	// - "Show text labels" (mode 0): no TBSTYLE_LIST (text below icons)
	// - "Selective text on right" (mode 1): TBSTYLE_LIST (text beside icons)
	// - "No text labels" (mode 2): TBSTYLE_LIST doesn't matter, but keep it off
	DWORD dwTbStyle = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS |
		TBSTYLE_FLAT | TBSTYLE_TOOLTIPS |
		CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE | CCS_ADJUSTABLE;

	if (dwTextMode == CE_TEXTMODE_SELECTIVE)
		dwTbStyle |= TBSTYLE_LIST;

	// Create the toolbar control
	m_hWndToolbar = CreateWindowEx(
		0,
		TOOLBARCLASSNAME,
		nullptr,
		dwTbStyle,
		0, 0, 0, 0,
		hWndParent,
		nullptr,
		_AtlBaseModule.GetModuleInstance(),
		nullptr);

	if (!m_hWndToolbar)
		return E_FAIL;

	// Required: set the struct size for the toolbar
	SendMessage(m_hWndToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

	// Extended styles: draw dropdown arrows
	DWORD dwExStyle = TBSTYLE_EX_DRAWDDARROWS;
	if (dwTextMode == CE_TEXTMODE_SELECTIVE)
		dwExStyle |= TBSTYLE_EX_MIXEDBUTTONS;
	SendMessage(m_hWndToolbar, TB_SETEXTENDEDSTYLE, 0, dwExStyle);

	// Text rows: "Show text labels" = 1 row, "No text labels" = 0 rows,
	// "Selective text" = doesn't matter (MIXEDBUTTONS handles it)
	if (dwTextMode == CE_TEXTMODE_SHOWTEXT)
		SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 1, 0);
	else if (dwTextMode == CE_TEXTMODE_NOTEXT)
		SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 0, 0);

	// Remember the active theme so we can detect switches later
	m_lastTheme = settings.theme;

	// Load toolbar bitmaps — select size based on settings, theme-aware
	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	COLORREF crMask = RGB(255, 0, 255);

	bool fIE55 = (settings.theme == CLASSIC_EXPLORER_2K && settings.ie55Style == 1);
	TB_SKIN_BITMAPS bmp = GetActiveBitmapIDs(settings.theme, fSmallIcons, fIE55);

	m_hilDefault = _CreateImageListFromResource(hInst, bmp.idDef, bmp.cx, crMask);
	m_hilHot     = _CreateImageListFromResource(hInst, bmp.idHot, bmp.cx, crMask);

	// Create disabled image list (algorithm differs per theme)
	m_hilDisabled = _CreateDisabledImageList(hInst, bmp.idDef, bmp.cx, crMask, settings.theme);

	// Append system bitmap icons for Properties, Cut, Copy, Paste, Folder Options
	_AppendSystemBitmapIcons(m_hilDefault, m_hilHot, m_hilDisabled, bmp.cx, settings.theme, fIE55);

	SendMessage(m_hWndToolbar, TB_SETIMAGELIST, 0, (LPARAM)m_hilDefault);
	SendMessage(m_hWndToolbar, TB_SETHOTIMAGELIST, 0, (LPARAM)m_hilHot);
	if (m_hilDisabled)
		SendMessage(m_hWndToolbar, TB_SETDISABLEDIMAGELIST, 0, (LPARAM)m_hilDisabled);

	// Pre-load all button strings into the toolbar's string pool
	PreloadAllStrings();

	// Populate buttons (from saved layout or defaults)
	InitToolbarButtons();

	// Autosize the toolbar
	SendMessage(m_hWndToolbar, TB_AUTOSIZE, 0, 0);

	// Install a subclass for message handling
	SetWindowSubclass(m_hWndToolbar, ToolbarSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));

	// Install rebar subclass to intercept WM_COMMAND and WM_NOTIFY from our toolbar
	InstallRebarHook();

	return S_OK;
}

void CStandardToolbar::InitToolbarButtons()
{
	if (!m_hWndToolbar)
		return;

	// Get the active skin's catalog and default layout
	CEUtil::CESettings settings = CEUtil::GetCESettings();
	int nCatalog = 0;
	const TB_BUTTON_DEF* pCatalog = GetActiveButtonCatalog(settings.theme, &nCatalog);
	int nDefLayout = 0;
	const int* pDefLayout = GetActiveDefaultLayout(settings.theme, &nDefLayout);

	// Try to load saved layout from registry, fall back to defaults
	int layout[MAX_SAVED_BUTTONS];
	int nLayout = 0;
	if (!LoadToolbarLayout(layout, &nLayout))
	{
		memcpy(layout, pDefLayout, nDefLayout * sizeof(int));
		nLayout = nDefLayout;
	}

	// Build TBBUTTON array from layout
	TBBUTTON tbb[MAX_SAVED_BUTTONS] = {};
	int nButtons = 0;

	for (int i = 0; i < nLayout && nButtons < MAX_SAVED_BUTTONS; i++)
	{
		if (layout[i] == 0)
		{
			// Separator
			tbb[nButtons].iBitmap = 0;
			tbb[nButtons].idCommand = 0;
			tbb[nButtons].fsState = 0;
			tbb[nButtons].fsStyle = BTNS_SEP;
			tbb[nButtons].dwData = 0;
			tbb[nButtons].iString = -1;
		}
		else
		{
			// Find button definition in catalog
			int idx = FindButtonCatalogIndex(pCatalog, nCatalog, layout[i]);
			if (idx < 0) continue;

			const TB_BUTTON_DEF& def = pCatalog[idx];
			tbb[nButtons].iBitmap = def.iBitmap;
			tbb[nButtons].idCommand = def.idCommand;
			tbb[nButtons].fsState = TBSTATE_ENABLED;
			tbb[nButtons].fsStyle = def.fsStyle;
			tbb[nButtons].dwData = 0;
			tbb[nButtons].iString = m_iStringPool[idx];

			// IE 5.5 style: add text to History button on 2K skin
			if (settings.theme == CLASSIC_EXPLORER_2K && settings.ie55Style == 1
				&& def.idCommand == TBIDM_HISTORY)
			{
				tbb[nButtons].fsStyle |= BTNS_SHOWTEXT;
			}

			// Win98 Views: make Views button a split dropdown (click cycles, arrow drops down)
			if (settings.win98Views == 1 && def.idCommand == TBIDM_VIEWMENU)
			{
				tbb[nButtons].fsStyle = (tbb[nButtons].fsStyle & ~BTNS_WHOLEDROPDOWN) | BTNS_DROPDOWN;
			}
		}
		nButtons++;
	}

	SendMessage(m_hWndToolbar, TB_ADDBUTTONS, nButtons, (LPARAM)tbb);
}

void CStandardToolbar::DestroyToolbarWindow()
{
	if (m_hilDefault)
	{
		ImageList_Destroy(m_hilDefault);
		m_hilDefault = nullptr;
	}
	if (m_hilHot)
	{
		ImageList_Destroy(m_hilHot);
		m_hilHot = nullptr;
	}
	if (m_hilDisabled)
	{
		ImageList_Destroy(m_hilDisabled);
		m_hilDisabled = nullptr;
	}
	if (m_hWndToolbar && IsWindow(m_hWndToolbar))
	{
		DestroyWindow(m_hWndToolbar);
		m_hWndToolbar = nullptr;
	}
}

// ================================================================================================
// Pre-load all button strings into the toolbar's string pool.
// Must be called before InitToolbarButtons() so that string indices are available.
//

void CStandardToolbar::PreloadAllStrings()
{
	memset(m_iStringPool, -1, sizeof(m_iStringPool));

	int nCatalog = 0;
	const TB_BUTTON_DEF* pCatalog = GetActiveButtonCatalog(m_lastTheme, &nCatalog);

	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	for (int i = 0; i < nCatalog; i++)
	{
		WCHAR szText[64] = {};
		LoadStringW(hInst, pCatalog[i].idsString, szText, ARRAYSIZE(szText));
		m_iStringPool[i] = (int)SendMessage(m_hWndToolbar, TB_ADDSTRING, 0, (LPARAM)szText);
	}
}

// ================================================================================================
// Save / Load toolbar layout to/from registry.
// Layout is stored as a REG_BINARY array of DWORDs (command IDs, 0 = separator).
//

void CStandardToolbar::SaveToolbarLayout()
{
	if (!m_hWndToolbar)
		return;

	int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
	if (nButtons <= 0 || nButtons > MAX_SAVED_BUTTONS)
		return;

	DWORD layout[MAX_SAVED_BUTTONS] = {};
	for (int i = 0; i < nButtons; i++)
	{
		TBBUTTON tbb = {};
		SendMessage(m_hWndToolbar, TB_GETBUTTON, i, (LPARAM)&tbb);
		layout[i] = (tbb.fsStyle & BTNS_SEP) ? 0 : (DWORD)tbb.idCommand;
	}

	HKEY hKey;
	if (RegCreateKeyExW(HKEY_CURRENT_USER, L"SOFTWARE\\kawapure\\ClassicExplorer",
		0, nullptr, 0, KEY_WRITE, nullptr, &hKey, nullptr) == ERROR_SUCCESS)
	{
		RegSetValueExW(hKey, L"ToolbarLayout", 0, REG_BINARY,
			(const BYTE*)layout, nButtons * sizeof(DWORD));
		RegCloseKey(hKey);
	}
}

bool CStandardToolbar::LoadToolbarLayout(int* pLayout, int* pCount)
{
	*pCount = 0;

	HKEY hKey;
	if (RegOpenKeyExW(HKEY_CURRENT_USER, L"SOFTWARE\\kawapure\\ClassicExplorer",
		0, KEY_READ, &hKey) != ERROR_SUCCESS)
		return false;

	DWORD cbData = MAX_SAVED_BUTTONS * sizeof(DWORD);
	DWORD buf[MAX_SAVED_BUTTONS] = {};
	DWORD dwType = 0;

	LSTATUS ls = RegQueryValueExW(hKey, L"ToolbarLayout", nullptr, &dwType,
		(BYTE*)buf, &cbData);
	RegCloseKey(hKey);

	if (ls != ERROR_SUCCESS || dwType != REG_BINARY || cbData < sizeof(DWORD))
		return false;

	int n = (int)(cbData / sizeof(DWORD));
	if (n > MAX_SAVED_BUTTONS)
		n = MAX_SAVED_BUTTONS;

	for (int i = 0; i < n; i++)
		pLayout[i] = (int)buf[i];

	*pCount = n;
	return true;
}

void CStandardToolbar::DeleteSavedToolbarLayout()
{
	HKEY hKey;
	if (RegOpenKeyExW(HKEY_CURRENT_USER, L"SOFTWARE\\kawapure\\ClassicExplorer",
		0, KEY_WRITE, &hKey) == ERROR_SUCCESS)
	{
		RegDeleteValueW(hKey, L"ToolbarLayout");
		RegCloseKey(hKey);
	}
}

// ================================================================================================
// TB_CUSTOMIZE notification handlers
// Ported from Windows XP browseui itbar.cpp TOOLSBANDCLASS.
//

void CStandardToolbar::OnBeginCustomize(HWND hCustDlg)
{
	// Check if the child dialog already exists (TBN_INITCUSTOMIZE is re-sent on Reset)
	if (m_hwndCustomizeChild && IsWindow(m_hwndCustomizeChild))
		return;

	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	m_hwndCustomizeChild = CreateDialogW(hInst,
		MAKEINTRESOURCEW(IDD_CUSTOMIZE_TOOLBAR), hCustDlg, CustomizeChildDlgProc);

	if (!m_hwndCustomizeChild)
		return;

	// Enlarge the comctl32 customize dialog to make room for our child dialog
	RECT rcCustWnd, rcCustClient, rcChild;
	GetWindowRect(hCustDlg, &rcCustWnd);
	GetClientRect(hCustDlg, &rcCustClient);
	GetWindowRect(m_hwndCustomizeChild, &rcChild);

	int childHeight = rcChild.bottom - rcChild.top;
	int custWidth = rcCustWnd.right - rcCustWnd.left;
	int custHeight = rcCustWnd.bottom - rcCustWnd.top;

	SetWindowPos(hCustDlg, nullptr,
		rcCustWnd.left, rcCustWnd.top,
		custWidth, custHeight + childHeight,
		SWP_NOZORDER);

	// Position our child dialog at the bottom of the original client area
	SetWindowPos(m_hwndCustomizeChild, HWND_TOP,
		rcCustClient.left, rcCustClient.bottom,
		0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
}

void CStandardToolbar::OnEndCustomize()
{
	SaveToolbarLayout();
	m_hwndCustomizeChild = nullptr;
	UpdateButtonStates();

	// Notify the rebar parent to recalculate band sizes
	HWND hRebar = GetParent(m_hWndToolbar);
	if (hRebar && IsWindow(hRebar))
	{
		int nBands = (int)SendMessage(hRebar, RB_GETBANDCOUNT, 0, 0);
		for (int i = 0; i < nBands; i++)
		{
			REBARBANDINFO rbi = { sizeof(rbi) };
			rbi.fMask = RBBIM_CHILD;
			SendMessage(hRebar, RB_GETBANDINFO, i, (LPARAM)&rbi);
			if (rbi.hwndChild == m_hWndToolbar)
			{
				DESKBANDINFO dbi = {};
				dbi.dwMask = DBIM_MINSIZE | DBIM_ACTUAL;
				GetBandInfo(0, 0, &dbi);

				rbi.fMask = RBBIM_CHILDSIZE;
				rbi.cxMinChild = dbi.ptMinSize.x;
				rbi.cyMinChild = dbi.ptMinSize.y;
				SendMessage(hRebar, RB_SETBANDINFO, i, (LPARAM)&rbi);
				break;
			}
		}
	}
}

LRESULT CStandardToolbar::OnGetButtonInfo(LPNMTOOLBAR ptbn)
{
	int nCatalog = 0;
	const TB_BUTTON_DEF* pCatalog = GetActiveButtonCatalog(m_lastTheme, &nCatalog);

	if (ptbn->iItem >= 0 && ptbn->iItem < nCatalog)
	{
		const TB_BUTTON_DEF& def = pCatalog[ptbn->iItem];

		ptbn->tbButton.iBitmap = def.iBitmap;
		ptbn->tbButton.idCommand = def.idCommand;
		ptbn->tbButton.fsState = TBSTATE_ENABLED;
		ptbn->tbButton.fsStyle = def.fsStyle;
		ptbn->tbButton.dwData = 0;
		ptbn->tbButton.iString = m_iStringPool[ptbn->iItem];

		// Also fill pszText for the customize dialog's "Available buttons" list
		if (ptbn->pszText && ptbn->cchText > 0)
		{
			WCHAR szText[64] = {};
			LoadStringW(_AtlBaseModule.GetModuleInstance(), def.idsString, szText, ARRAYSIZE(szText));
			wcsncpy_s(ptbn->pszText, ptbn->cchText, szText, _TRUNCATE);
		}

		return TRUE;
	}
	return FALSE;
}

LRESULT CStandardToolbar::OnReset()
{
	// Delete saved layout from registry
	DeleteSavedToolbarLayout();

	// Get the current skin's defaults
	TB_SKIN_DEFAULTS skinDef = GetDefaultSettings(m_lastTheme);
	int nCatalog = 0;
	const TB_BUTTON_DEF* pCatalog = GetActiveButtonCatalog(m_lastTheme, &nCatalog);
	int nDefLayout = 0;
	const int* pDefLayout = GetActiveDefaultLayout(m_lastTheme, &nDefLayout);

	// Reset text/icon settings to the current skin's defaults
	CEUtil::CESettings resetSettings(CLASSIC_EXPLORER_NONE, -1, -1, -1,
		skinDef.smallIcons, skinDef.textLabelMode, 0);
	CEUtil::WriteCESettings(resetSettings);

	// Update child dialog combo selections to reflect defaults
	if (m_hwndCustomizeChild && IsWindow(m_hwndCustomizeChild))
	{
		_SetComboSelection(GetDlgItem(m_hwndCustomizeChild, IDC_SHOWTEXT),
			_TextModeToIds(skinDef.textLabelMode));
		_SetComboSelection(GetDlgItem(m_hwndCustomizeChild, IDC_SMALLICONS),
			skinDef.smallIcons ? IDS_SMALLICONS : IDS_LARGEICONS);
		SendMessage(GetDlgItem(m_hwndCustomizeChild, IDC_IE55STYLE), BM_SETCHECK, BST_UNCHECKED, 0);
		SendMessage(GetDlgItem(m_hwndCustomizeChild, IDC_WIN98VIEWS), BM_SETCHECK, BST_UNCHECKED, 0);
	}

	// Remove all buttons from the toolbar
	int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
	while (nButtons-- > 0)
		SendMessage(m_hWndToolbar, TB_DELETEBUTTON, 0, 0);

	// Re-add default buttons for this skin
	TBBUTTON tbb[MAX_SAVED_BUTTONS] = {};
	int nDefButtons = 0;

	for (int i = 0; i < nDefLayout && nDefButtons < MAX_SAVED_BUTTONS; i++)
	{
		if (pDefLayout[i] == 0)
		{
			tbb[nDefButtons].fsStyle = BTNS_SEP;
			tbb[nDefButtons].iString = -1;
		}
		else
		{
			int idx = FindButtonCatalogIndex(pCatalog, nCatalog, pDefLayout[i]);
			if (idx < 0) continue;

			const TB_BUTTON_DEF& def = pCatalog[idx];
			tbb[nDefButtons].iBitmap = def.iBitmap;
			tbb[nDefButtons].idCommand = def.idCommand;
			tbb[nDefButtons].fsState = TBSTATE_ENABLED;
			tbb[nDefButtons].fsStyle = def.fsStyle;
			tbb[nDefButtons].iString = m_iStringPool[idx];
		}
		nDefButtons++;
	}

	SendMessage(m_hWndToolbar, TB_ADDBUTTONS, nDefButtons, (LPARAM)tbb);

	return 0;  // Let comctl32 refresh the customize dialog lists
}

// ================================================================================================
// Button state management
//

void CStandardToolbar::UpdateButtonStates()
{
	if (!m_hWndToolbar)
		return;

	// Back / Forward: use cached state from OnCommandStateChange
	SendMessage(m_hWndToolbar, TB_ENABLEBUTTON, TBIDM_BACK, MAKELONG(m_bCanGoBack, 0));
	SendMessage(m_hWndToolbar, TB_ENABLEBUTTON, TBIDM_FORWARD, MAKELONG(m_bCanGoForward, 0));

	// Up: check if we can go up (the shell view has a parent folder)
	BOOL bCanGoUp = FALSE;
	if (m_pShellBrowser)
	{
		// Try to get current folder and see if it has a parent
		CComPtr<IShellView> pView;
		if (SUCCEEDED(m_pShellBrowser->QueryActiveShellView(&pView)))
		{
			CComQIPtr<IFolderView> pFolderView = pView;
			if (pFolderView)
			{
				CComPtr<IPersistFolder2> pFolder;
				if (SUCCEEDED(pFolderView->GetFolder(IID_IPersistFolder2, (void**)&pFolder)))
				{
					PIDLIST_ABSOLUTE pidl = nullptr;
					if (SUCCEEDED(pFolder->GetCurFolder(&pidl)) && pidl)
					{
						// Desktop has an empty PIDL (size == 2), anything else can go up
						if (ILGetSize(pidl) > 2)
							bCanGoUp = TRUE;
						CoTaskMemFree(pidl);
					}
				}
			}
		}
	}
	SendMessage(m_hWndToolbar, TB_ENABLEBUTTON, TBIDM_PREVIOUSFOLDER, MAKELONG(bCanGoUp, 0));

	// Folders: sync the check state from the property bag.
	// The IPropertyBag::Write vtable hook (BagWriteHook) handles reactive
	// updates — this path covers the initial call and navigation events.
	if (!m_pBrowserBag)
		GetBrowserBag();

	SyncFoldersCheckState();
}

void CStandardToolbar::SyncFoldersCheckState()
{
	if (!m_hWndToolbar || !m_pBrowserBag)
		return;

	VARIANT val = {};
	val.vt = VT_EMPTY;
	bool bNavPane = SUCCEEDED(m_pBrowserBag->Read(g_NavPaneVisible, &val, nullptr))
		&& val.vt == VT_BOOL && val.boolVal;
	VariantClear(&val);
	SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_ALLFOLDERS, MAKELONG(bNavPane, 0));
}

// ================================================================================================
// Command routing
//
// We subclass the rebar (toolbar's parent) to intercept WM_COMMAND and WM_NOTIFY
// from our toolbar control, then dispatch to our HandleCommand / dropdown handlers.

static LRESULT CALLBACK RebarSubclassProc(
	HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
	UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	CStandardToolbar* pThis = reinterpret_cast<CStandardToolbar*>(dwRefData);

	switch (uMsg)
	{
	case WM_COMMAND:
	{
		int idCmd = LOWORD(wParam);
		// Check if this command is from our toolbar
		if ((HWND)lParam == pThis->GetToolbarHwnd())
		{
			pThis->HandleCommand(idCmd);
			return 0;
		}
		break;
	}

	case WM_NOTIFY:
	{
		NMHDR* pnmh = reinterpret_cast<NMHDR*>(lParam);
		if (pnmh && pnmh->hwndFrom == pThis->GetToolbarHwnd())
		{
			switch (pnmh->code)
			{
			case TBN_DROPDOWN:
			{
				NMTOOLBAR* pnmtb = reinterpret_cast<NMTOOLBAR*>(lParam);
				int idCmd = pnmtb->iItem;

				if (idCmd == TBIDM_BACK || idCmd == TBIDM_FORWARD)
				{
					pThis->ShowBackForwardMenu(idCmd, pnmtb);
					return TBDDRET_DEFAULT;
				}
				else if (idCmd == TBIDM_VIEWMENU)
				{
					pThis->OnViewsDropdown(pnmtb);
					return TBDDRET_DEFAULT;
				}
				break;
			}

			case TBN_INITCUSTOMIZE:
			{
				NMTBCUSTOMIZEDLG* pnm = reinterpret_cast<NMTBCUSTOMIZEDLG*>(lParam);
				if (pnm->hDlg && IsWindow(pnm->hDlg))
					pThis->OnBeginCustomize(pnm->hDlg);
				return TBNRF_HIDEHELP;
			}

			case TBN_GETBUTTONINFOW:
			{
				NMTOOLBAR* ptbn = reinterpret_cast<NMTOOLBAR*>(lParam);
				return pThis->OnGetButtonInfo(ptbn);
			}

			case TBN_QUERYINSERT:
				return TRUE;

			case TBN_QUERYDELETE:
				return TRUE;

			case TBN_RESET:
				return pThis->OnReset();

			case TBN_ENDADJUST:
				pThis->OnEndCustomize();
				break;

			case TBN_TOOLBARCHANGE:
				break;

			case TBN_GETINFOTIP:
			{
				// Tooltips are handled via TB_ADDSTRING
				break;
			}
			}
		}
		break;
	}

	case WM_NCDESTROY:
		RemoveWindowSubclass(hWnd, RebarSubclassProc, uIdSubclass);
		break;
	}

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// Public accessors used by the subclass proc
HWND CStandardToolbar::GetToolbarHwnd() const
{
	return m_hWndToolbar;
}

void CStandardToolbar::HandleCommand(int idCmd)
{
	switch (idCmd)
	{
	case TBIDM_BACK:           OnBack(); break;
	case TBIDM_FORWARD:        OnForward(); break;
	case TBIDM_PREVIOUSFOLDER: OnUpOneLevel(); break;
	case TBIDM_SEARCH:         OnSearch(); break;
	case TBIDM_ALLFOLDERS:     OnFolders(); break;
	case TBIDM_VIEWMENU:       CycleViewMode(); break;
	case TBIDM_MOVETO:         OnMoveTo(); break;
	case TBIDM_COPYTO:         OnCopyTo(); break;
	case TBIDM_DELETE:          OnDelete(); break;
	case TBIDM_UNDO:           OnUndo(); break;
	case TBIDM_STOPDOWNLOAD:   OnStop(); break;
	case TBIDM_REFRESH:        OnRefresh(); break;
	case TBIDM_HOME:           OnHome(); break;
	case TBIDM_CONNECT:        OnMapNetDrive(); break;
	case TBIDM_DISCONNECT:     OnDisconnectNetDrive(); break;
	case TBIDM_FAVORITES:      OnFavorites(); break;
	case TBIDM_HISTORY:        OnHistory(); break;
	case TBIDM_THEATER:        OnFullScreen(); break;
	case TBIDM_PROPERTIES:     OnProperties(); break;
	case TBIDM_CUT:            OnCut(); break;
	case TBIDM_COPY:           OnCopy(); break;
	case TBIDM_PASTE:          OnPaste(); break;
	case TBIDM_FOLDEROPTIONS:  OnFolderOptions(); break;
	}
}

// Install the rebar subclass (called from CreateToolbarWindow after toolbar is created)
void CStandardToolbar::InstallRebarHook()
{
	if (m_hWndToolbar)
	{
		HWND hRebar = GetParent(m_hWndToolbar);
		if (hRebar)
		{
			SetWindowSubclass(hRebar, RebarSubclassProc, (UINT_PTR)this,
				reinterpret_cast<DWORD_PTR>(this));
		}
	}
}

// ================================================================================================
// Navigation command handlers
//

void CStandardToolbar::OnBack()
{
	if (m_pShellBrowser)
		m_pShellBrowser->BrowseObject(nullptr, SBSP_NAVIGATEBACK | SBSP_SAMEBROWSER);
}

void CStandardToolbar::OnForward()
{
	if (m_pShellBrowser)
		m_pShellBrowser->BrowseObject(nullptr, SBSP_NAVIGATEFORWARD | SBSP_SAMEBROWSER);
}

void CStandardToolbar::OnUpOneLevel()
{
	if (m_pShellBrowser)
		m_pShellBrowser->BrowseObject(nullptr, SBSP_PARENT | SBSP_SAMEBROWSER);
}

void CStandardToolbar::OnSearch()
{
	// On Windows 10/11, focus the search box in the Explorer window.
	// The search box is a child of the ShellTabWindowClass.
	// Send Ctrl+E which activates the search box on modern Explorer.
	HWND hTop = m_hWndParent ? GetAncestor(m_hWndParent, GA_ROOT) : nullptr;
	if (hTop)
	{
		// Simulate Ctrl+E keystroke to activate the search box
		INPUT inputs[4] = {};
		inputs[0].type = INPUT_KEYBOARD;
		inputs[0].ki.wVk = VK_CONTROL;
		inputs[1].type = INPUT_KEYBOARD;
		inputs[1].ki.wVk = 'E';
		inputs[2].type = INPUT_KEYBOARD;
		inputs[2].ki.wVk = 'E';
		inputs[2].ki.dwFlags = KEYEVENTF_KEYUP;
		inputs[3].type = INPUT_KEYBOARD;
		inputs[3].ki.wVk = VK_CONTROL;
		inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;
		SetForegroundWindow(hTop);
		SendInput(ARRAYSIZE(inputs), inputs, sizeof(INPUT));
	}
}

void CStandardToolbar::OnFolders()
{
	// Toggle the navigation pane using the per-browser property bag
	// (same approach as Open-Shell)
	if (!m_pBrowserBag)
		GetBrowserBag();

	if (m_pBrowserBag)
	{
		VARIANT val = {};
		val.vt = VT_EMPTY;
		bool bNavPane = SUCCEEDED(m_pBrowserBag->Read(g_NavPaneVisible, &val, nullptr))
			&& val.vt == VT_BOOL && val.boolVal;
		VariantClear(&val);

		bool bNewState = !bNavPane;
		val.vt = VT_BOOL;
		val.boolVal = bNewState ? VARIANT_TRUE : VARIANT_FALSE;
		m_pBrowserBag->Write(g_NavPaneVisible, &val);

		// Explicitly sync the check state to match what we wrote.
		// BTNS_CHECK auto-toggles on click, which can desync if the
		// initial state was wrong (e.g. Windhawk exploring mode sets
		// the nav pane visible after our first UpdateButtonStates).
		SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_ALLFOLDERS, MAKELONG(bNewState, 0));
	}
}

void CStandardToolbar::OnMoveTo()
{
	SendDefViewCommand(DEFVIEW_MOVETO);
}

void CStandardToolbar::OnCopyTo()
{
	SendDefViewCommand(DEFVIEW_COPYTO);
}

void CStandardToolbar::OnDelete()
{
	SendDefViewCommand(DEFVIEW_DELETE);
}

void CStandardToolbar::OnUndo()
{
	SendShellTabCommand(SHELLTAB_UNDO);
}

void CStandardToolbar::OnStop()
{
	if (m_pWebBrowser)
		m_pWebBrowser->Stop();
}

void CStandardToolbar::OnRefresh()
{
	SendShellTabCommand(SHELLTAB_REFRESH);
}

void CStandardToolbar::OnHome()
{
	if (m_pWebBrowser)
		m_pWebBrowser->GoHome();
}

void CStandardToolbar::OnMapNetDrive()
{
	SendShellTabCommand(SHELLTAB_MAPDRIVE);
}

void CStandardToolbar::OnDisconnectNetDrive()
{
	SendShellTabCommand(SHELLTAB_DISCONNECT);
}

void CStandardToolbar::OnFavorites()
{
	// Open the Favorites folder in a new Explorer window
	if (m_pShellBrowser)
	{
		PIDLIST_ABSOLUTE pidl = nullptr;
		if (SUCCEEDED(SHGetSpecialFolderLocation(nullptr, CSIDL_FAVORITES, &pidl)) && pidl)
		{
			m_pShellBrowser->BrowseObject(pidl, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
			CoTaskMemFree(pidl);
		}
	}
}

void CStandardToolbar::OnHistory()
{
	// Navigate to the History special folder
	if (m_pShellBrowser)
	{
		PIDLIST_ABSOLUTE pidl = nullptr;
		if (SUCCEEDED(SHGetSpecialFolderLocation(nullptr, CSIDL_HISTORY, &pidl)) && pidl)
		{
			m_pShellBrowser->BrowseObject(pidl, SBSP_SAMEBROWSER | SBSP_ABSOLUTE);
			CoTaskMemFree(pidl);
		}
	}
}

void CStandardToolbar::OnFullScreen()
{
	// Toggle full-screen / theater mode
	if (m_pWebBrowser)
	{
		VARIANT_BOOL bFullScreen = VARIANT_FALSE;
		m_pWebBrowser->get_TheaterMode(&bFullScreen);
		m_pWebBrowser->put_TheaterMode(bFullScreen == VARIANT_FALSE ? VARIANT_TRUE : VARIANT_FALSE);
	}
}

void CStandardToolbar::OnProperties()
{
	SendDefViewCommand(DEFVIEW_PROPERTIES);
}

void CStandardToolbar::OnCut()
{
	SendDefViewCommand(DEFVIEW_CUT);
}

void CStandardToolbar::OnCopy()
{
	SendDefViewCommand(DEFVIEW_COPY);
}

void CStandardToolbar::OnPaste()
{
	SendDefViewCommand(DEFVIEW_PASTE);
}

void CStandardToolbar::OnFolderOptions()
{
	// Launch Folder Options dialog directly (the SFVIDM_TOOL_OPTIONS command
	// isn't handled by modern Explorer's ShellTabWindowClass)
	ShellExecuteW(nullptr, nullptr, L"rundll32.exe",
		L"shell32.dll,Options_RunDLL 0", nullptr, SW_SHOWNORMAL);
}

// ================================================================================================
// Views dropdown menu
//

void CStandardToolbar::OnViewsDropdown(LPNMTOOLBAR pnmtb)
{
	if (!m_pShellBrowser || !m_hWndToolbar)
		return;

	CEUtil::CESettings settings = CEUtil::GetCESettings();
	bool f2KViews = (settings.theme == CLASSIC_EXPLORER_2K);

	// Map current view to DefView command ID for radio check.
	// On Win10/11, both Thumbnails (28751) and Icons (28750) report FVM_ICON;
	// we use GetViewModeAndIconSize to tell them apart by icon pixel size.
	UINT uCurrentCmd = 0;
	CComPtr<IShellView> pView;
	if (SUCCEEDED(m_pShellBrowser->QueryActiveShellView(&pView)))
	{
		CComQIPtr<IFolderView2> pFV2 = pView;
		if (pFV2)
		{
			FOLDERVIEWMODE fvm = FVM_AUTO;
			int iIconSize = 0;
			if (SUCCEEDED(pFV2->GetViewModeAndIconSize(&fvm, &iIconSize)))
			{
				if (f2KViews)
				{
					// Shell converts FVM_ICON/32px to FVM_SMALLICON internally,
					// so disambiguate by icon size: >=32 = Large Icons, <32 = Small Icons
					switch (fvm)
					{
					case FVM_ICON:      uCurrentCmd = FVM_ICON;      break;
					case FVM_SMALLICON: uCurrentCmd = (iIconSize >= 32) ? FVM_ICON : FVM_SMALLICON; break;
					case FVM_LIST:      uCurrentCmd = FVM_LIST;       break;
					case FVM_DETAILS:   uCurrentCmd = FVM_DETAILS;    break;
					default:            uCurrentCmd = FVM_ICON;        break;
					}
				}
				else
				{
					switch (fvm)
					{
					case FVM_THUMBNAIL: uCurrentCmd = DEFVIEW_VIEW_THUMBNAILS; break;
					case FVM_TILE:      uCurrentCmd = DEFVIEW_VIEW_TILES;      break;
					case FVM_LIST:      uCurrentCmd = DEFVIEW_VIEW_LIST;       break;
					case FVM_DETAILS:   uCurrentCmd = DEFVIEW_VIEW_DETAILS;    break;
					case FVM_ICON:
						// Large icons (>48px) = Thumbnails, medium/small = Icons
						uCurrentCmd = (iIconSize > 48) ? DEFVIEW_VIEW_THUMBNAILS : DEFVIEW_VIEW_ICONS;
						break;
					default:
						uCurrentCmd = DEFVIEW_VIEW_ICONS;
						break;
					}
				}
			}
		}
	}

	// Build popup menu matching shell32's View submenu layout
	HMENU hMenu = CreatePopupMenu();
	if (!hMenu)
		return;

	if (f2KViews)
	{
		// Win2K: Large Icons, Small Icons, List, Details
		// Use FVM_* enum values as menu command IDs (1-4)
		AppendMenu(hMenu, MF_STRING, FVM_ICON,      L"Lar&ge Icons");
		AppendMenu(hMenu, MF_STRING, FVM_SMALLICON,  L"S&mall Icons");
		AppendMenu(hMenu, MF_STRING, FVM_LIST,       L"&List");
		AppendMenu(hMenu, MF_STRING, FVM_DETAILS,    L"&Details");

		if (uCurrentCmd)
			CheckMenuRadioItem(hMenu, FVM_ICON, FVM_DETAILS,
				uCurrentCmd, MF_BYCOMMAND);
	}
	else
	{
		// XP: Thumbnails, Tiles, Icons, List, Details
		AppendMenu(hMenu, MF_STRING, DEFVIEW_VIEW_THUMBNAILS, L"T&humbnails");
		AppendMenu(hMenu, MF_STRING, DEFVIEW_VIEW_TILES,      L"Tile&s");
		AppendMenu(hMenu, MF_STRING, DEFVIEW_VIEW_ICONS,      L"Ico&ns");
		AppendMenu(hMenu, MF_STRING, DEFVIEW_VIEW_LIST,       L"&List");
		AppendMenu(hMenu, MF_STRING, DEFVIEW_VIEW_DETAILS,    L"&Details");

		if (uCurrentCmd)
			CheckMenuRadioItem(hMenu, DEFVIEW_VIEW_DETAILS, DEFVIEW_VIEW_LIST,
				uCurrentCmd, MF_BYCOMMAND);
	}

	// Get button rect for positioning
	RECT rcButton = {};
	SendMessage(m_hWndToolbar, TB_GETRECT, TBIDM_VIEWMENU, (LPARAM)&rcButton);
	MapWindowPoints(m_hWndToolbar, HWND_DESKTOP, (LPPOINT)&rcButton, 2);

	UINT uCmd = TrackPopupMenuEx(
		hMenu,
		TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN,
		rcButton.left, rcButton.bottom,
		m_hWndToolbar,
		nullptr);

	DestroyMenu(hMenu);

	if (uCmd > 0)
	{
		if (f2KViews)
			SetClassicViewMode((FOLDERVIEWMODE)uCmd);
		else if (uCmd == DEFVIEW_VIEW_ICONS)
			SetClassicViewMode(FVM_ICON);  // 32px icons (pre-Vista)
		else
			SendDefViewCommand(uCmd);
	}
}

void CStandardToolbar::CycleViewMode()
{
	// When the button itself is clicked (not the dropdown), cycle through view modes.
	// 98 style: Large Icons -> Small Icons -> List -> Details -> Large Icons
	// XP style: Thumbnails -> Tiles -> Icons -> List -> Details -> Thumbnails
	if (!m_pShellBrowser)
		return;

	CComPtr<IShellView> pView;
	if (FAILED(m_pShellBrowser->QueryActiveShellView(&pView)))
		return;

	CComQIPtr<IFolderView2> pFV2 = pView;
	if (!pFV2)
		return;

	FOLDERVIEWMODE fvm = FVM_AUTO;
	int iIconSize = 0;
	if (FAILED(pFV2->GetViewModeAndIconSize(&fvm, &iIconSize)))
		return;

	CEUtil::CESettings settings = CEUtil::GetCESettings();

	if (settings.win98Views == 1)
	{
		// 98 cycle: Large Icons -> Small Icons -> List -> Details -> Large Icons
		// Shell converts FVM_ICON/32px to FVM_SMALLICON, so disambiguate by icon size
		FOLDERVIEWMODE nextFvm;
		bool isLargeIcons = (fvm == FVM_ICON) || (fvm == FVM_SMALLICON && iIconSize >= 32);
		if (isLargeIcons)
			nextFvm = FVM_SMALLICON;
		else switch (fvm)
		{
		case FVM_SMALLICON: nextFvm = FVM_LIST;      break;
		case FVM_LIST:      nextFvm = FVM_DETAILS;   break;
		case FVM_DETAILS:   nextFvm = FVM_ICON;      break;
		default:            nextFvm = FVM_ICON;      break;
		}
		SetClassicViewMode(nextFvm);
	}
	else
	{
		// XP cycle: Thumbnails -> Tiles -> Icons -> List -> Details -> Thumbnails
		// Disambiguate FVM_ICON by icon size (>48 = Thumbnails/Large, <=48 = Icons/Medium)
		int nextCmd;
		if (fvm == FVM_ICON && iIconSize > 48)
			nextCmd = DEFVIEW_VIEW_TILES;       // Thumbnails -> Tiles
		else if (fvm == FVM_THUMBNAIL)
			nextCmd = DEFVIEW_VIEW_TILES;       // Thumbnails -> Tiles
		else if (fvm == FVM_TILE)
		{
			SetClassicViewMode(FVM_ICON);       // Tiles -> Icons (32px)
			return;
		}
		else if (fvm == FVM_ICON)
			nextCmd = DEFVIEW_VIEW_LIST;        // Icons -> List
		else if (fvm == FVM_LIST)
			nextCmd = DEFVIEW_VIEW_DETAILS;     // List -> Details
		else if (fvm == FVM_DETAILS)
			nextCmd = DEFVIEW_VIEW_THUMBNAILS;  // Details -> Thumbnails
		else
			nextCmd = DEFVIEW_VIEW_DETAILS;     // Default to Details

		SendDefViewCommand(nextCmd);
	}
}

// ================================================================================================
// Back/Forward history dropdown
//

void CStandardToolbar::ShowBackForwardMenu(int idCmd, LPNMTOOLBAR pnmtb)
{
	if (!m_pWebBrowser)
		return;

	// Use ITravelLogStg to get the travel log entries
	CComQIPtr<IServiceProvider> pSP = m_pWebBrowser;
	if (!pSP)
		return;

	CComPtr<ITravelLogStg> pTravelLog;
	if (FAILED(pSP->QueryService(SID_STravelLogCursor, IID_ITravelLogStg, (void**)&pTravelLog)))
		return;

	// Enumerate entries
	DWORD dwFlags = (idCmd == TBIDM_BACK) ? TLEF_RELATIVE_BACK : TLEF_RELATIVE_FORE;

	CComPtr<IEnumTravelLogEntry> pEnum;
	if (FAILED(pTravelLog->EnumEntries(dwFlags, &pEnum)))
		return;

	HMENU hMenu = CreatePopupMenu();
	if (!hMenu)
		return;

	ITravelLogEntry* pEntry = nullptr;
	ULONG cFetched = 0;
	int nItem = 0;
	const int MAX_HISTORY_ITEMS = 9;

	while (nItem < MAX_HISTORY_ITEMS && pEnum->Next(1, &pEntry, &cFetched) == S_OK && cFetched > 0)
	{
		LPWSTR pwszTitle = nullptr;
		pEntry->GetTitle(&pwszTitle);

		WCHAR szDisplay[64] = {};
		if (pwszTitle && pwszTitle[0])
		{
			wcsncpy_s(szDisplay, pwszTitle, _TRUNCATE);
			CoTaskMemFree(pwszTitle);
		}
		else
		{
			// Fallback: use URL
			LPWSTR pwszUrl = nullptr;
			pEntry->GetURL(&pwszUrl);
			if (pwszUrl)
			{
				wcsncpy_s(szDisplay, pwszUrl, _TRUNCATE);
				CoTaskMemFree(pwszUrl);
			}
			else
			{
				wcscpy_s(szDisplay, L"(untitled)");
			}
		}

		nItem++;
		AppendMenu(hMenu, MF_STRING, nItem, szDisplay);
		pEntry->Release();
		pEntry = nullptr;
	}

	if (nItem == 0)
	{
		DestroyMenu(hMenu);
		return;
	}

	// Position the menu below the button
	RECT rcButton = {};
	SendMessage(m_hWndToolbar, TB_GETRECT, idCmd, (LPARAM)&rcButton);
	MapWindowPoints(m_hWndToolbar, HWND_DESKTOP, (LPPOINT)&rcButton, 2);

	UINT uCmd = TrackPopupMenuEx(
		hMenu,
		TPM_RETURNCMD | TPM_NONOTIFY | TPM_LEFTALIGN | TPM_TOPALIGN,
		rcButton.left, rcButton.bottom,
		m_hWndToolbar,
		nullptr);

	DestroyMenu(hMenu);

	if (uCmd > 0)
	{
		// Navigate back/forward by the selected number of steps
		CComQIPtr<IWebBrowser2> pWB = m_pWebBrowser;
		if (pWB)
		{
			// Re-enumerate to get the actual entry
			CComPtr<IEnumTravelLogEntry> pEnum2;
			if (SUCCEEDED(pTravelLog->EnumEntries(dwFlags, &pEnum2)))
			{
				ITravelLogEntry* pTarget = nullptr;
				ULONG cSkip = 0;
				for (ULONG i = 0; i < (ULONG)uCmd; i++)
				{
					if (pTarget) { pTarget->Release(); pTarget = nullptr; }
					pEnum2->Next(1, &pTarget, &cSkip);
				}
				if (pTarget)
				{
					pTravelLog->TravelTo(pTarget);
					pTarget->Release();
				}
			}
		}
	}
}

// ================================================================================================
// Helper: Send a WM_COMMAND to the ShellTabWindowClass parent
// (same pattern as Open-Shell's SendShellTabCommand)
//

void CStandardToolbar::SendShellTabCommand(int command)
{
	for (HWND hWnd = GetParent(m_hWndToolbar); hWnd; hWnd = GetParent(hWnd))
	{
		WCHAR szClass[64] = {};
		GetClassName(hWnd, szClass, ARRAYSIZE(szClass));
		if (_wcsicmp(szClass, L"ShellTabWindowClass") == 0)
		{
			SendMessage(hWnd, WM_COMMAND, command, 0);
			return;
		}
	}
}

// ================================================================================================
// Helper: Send a WM_COMMAND to the SHELLDLL_DefView window
// (same pattern as Open-Shell's defview command routing)
//

void CStandardToolbar::SendDefViewCommand(int command)
{
	if (!m_pShellBrowser)
		return;

	CComPtr<IShellView> pView;
	if (FAILED(m_pShellBrowser->QueryActiveShellView(&pView)))
		return;

	// Get the DefView window
	HWND hDefView = nullptr;
	pView->GetWindow(&hDefView);

	if (hDefView && IsWindow(hDefView))
		SendMessage(hDefView, WM_COMMAND, command, 0);
}

// Helper: Set view mode for classic (2K/98) view modes.
// Uses SendDefViewCommand for Small Icons (28752) since IFolderView2 doesn't
// handle FVM_SMALLICON properly, and SetViewModeAndIconSize for Large Icons (32px).
void CStandardToolbar::SetClassicViewMode(FOLDERVIEWMODE fvm)
{
	if (!m_pShellBrowser)
		return;

	switch (fvm)
	{
	case FVM_ICON:
	{
		// Large Icons: use IFolderView2 to set 32px icon size
		CComPtr<IShellView> pView;
		if (FAILED(m_pShellBrowser->QueryActiveShellView(&pView)))
			return;
		CComQIPtr<IFolderView2> pFV2 = pView;
		if (pFV2)
			pFV2->SetViewModeAndIconSize(FVM_ICON, 32);
		break;
	}
	case FVM_SMALLICON:
		SendDefViewCommand(DEFVIEW_VIEW_SMALLICONS);
		break;
	case FVM_LIST:
		SendDefViewCommand(DEFVIEW_VIEW_LIST);
		break;
	case FVM_DETAILS:
		SendDefViewCommand(DEFVIEW_VIEW_DETAILS);
		break;
	default:
		SendDefViewCommand(DEFVIEW_VIEW_ICONS);
		break;
	}
}

// ================================================================================================
// Helper: Obtain the per-browser property bag for nav pane toggling
// (same approach as Open-Shell: SID_FrameManager -> SID_PerBrowserPropertyBag)
//

void CStandardToolbar::GetBrowserBag()
{
	if (m_pBrowserBag)
		return;

	if (!m_pShellBrowser)
		return;

	CComPtr<IUnknown> pFrame;
	IUnknown_QueryService(m_pShellBrowser, SID_FrameManager, IID_IUnknown, (void**)&pFrame);
	if (pFrame)
		IUnknown_QueryService(pFrame, SID_PerBrowserPropertyBag, IID_IPropertyBag, (void**)&m_pBrowserBag);

	// Install the vtable hook on IPropertyBag::Write so we can reactively
	// detect when the nav pane visibility changes (same technique as Open-Shell).
	if (m_pBrowserBag)
	{
		g_hWndToolbarForHook = m_hWndToolbar;

		void** vtbl = *(void***)m_pBrowserBag.p;  // pointer to vtable
		void* pWrite = vtbl[4];                    // IPropertyBag::Write is slot 4 (QI, AddRef, Release, Read, Write)
		if (InterlockedCompareExchangePointer((void**)&g_OldBagWrite, pWrite, nullptr) == nullptr)
		{
			// First time — patch the vtable entry.
			DWORD oldProtect;
			VirtualProtect(&vtbl[4], sizeof(void*), PAGE_READWRITE, &oldProtect);
			vtbl[4] = (void*)BagWriteHook;
			VirtualProtect(&vtbl[4], sizeof(void*), oldProtect, &oldProtect);

			// Prevent our DLL from being unloaded while the vtable hook is live,
			// otherwise the hook would jump to freed memory.
			HMODULE hSelf;
			GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
				(LPCWSTR)BagWriteHook, &hSelf);
		}
	}
}

bool CStandardToolbar::EnsureBrowserBag()
{
	if (m_pBrowserBag)
		return false;
	GetBrowserBag();
	return m_pBrowserBag != nullptr;
}

// ================================================================================================
// Text label mode application
// Ported from Windows Server 2003 browseui _UpdateTextSettings / _UpdateToolsStyle.
//

void CStandardToolbar::ApplyTextLabelMode(DWORD dwMode)
{
	if (!m_hWndToolbar)
		return;

	switch (dwMode)
	{
	case CE_TEXTMODE_SHOWTEXT:
	{
		// Show text labels: no TBSTYLE_LIST, text rows = 1
		LONG style = GetWindowLong(m_hWndToolbar, GWL_STYLE);
		SetWindowLong(m_hWndToolbar, GWL_STYLE, style & ~TBSTYLE_LIST);
		SendMessage(m_hWndToolbar, TB_SETEXTENDEDSTYLE,
			TBSTYLE_EX_MIXEDBUTTONS, 0);
		SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 1, 0);
		break;
	}

	case CE_TEXTMODE_SELECTIVE:
	{
		// Selective text on right: TBSTYLE_LIST + MIXEDBUTTONS
		LONG style = GetWindowLong(m_hWndToolbar, GWL_STYLE);
		SetWindowLong(m_hWndToolbar, GWL_STYLE, style | TBSTYLE_LIST);
		SendMessage(m_hWndToolbar, TB_SETEXTENDEDSTYLE,
			TBSTYLE_EX_MIXEDBUTTONS, TBSTYLE_EX_MIXEDBUTTONS);
		SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 1, 0);
		break;
	}

	case CE_TEXTMODE_NOTEXT:
	{
		// No text labels: text rows = 0
		LONG style = GetWindowLong(m_hWndToolbar, GWL_STYLE);
		SetWindowLong(m_hWndToolbar, GWL_STYLE, style & ~TBSTYLE_LIST);
		SendMessage(m_hWndToolbar, TB_SETEXTENDEDSTYLE,
			TBSTYLE_EX_MIXEDBUTTONS, 0);
		SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 0, 0);
		break;
	}
	}

	// Force toolbar to recalculate layout
	SendMessage(m_hWndToolbar, TB_AUTOSIZE, 0, 0);
	InvalidateRect(m_hWndToolbar, nullptr, TRUE);
}

// ================================================================================================
// Rebuild image lists (called when icon size or theme changes).
//

void CStandardToolbar::RebuildImageLists()
{
	if (!m_hWndToolbar)
		return;

	CEUtil::CESettings settings = CEUtil::GetCESettings();
	bool fSmallIcons = (settings.smallIcons == 1);

	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	COLORREF crMask = RGB(255, 0, 255);

	bool fIE55 = (settings.theme == CLASSIC_EXPLORER_2K && settings.ie55Style == 1);
	TB_SKIN_BITMAPS bmp = GetActiveBitmapIDs(settings.theme, fSmallIcons, fIE55);

	// Destroy old image lists
	if (m_hilDefault) { ImageList_Destroy(m_hilDefault); m_hilDefault = nullptr; }
	if (m_hilHot)     { ImageList_Destroy(m_hilHot);     m_hilHot = nullptr; }
	if (m_hilDisabled){ ImageList_Destroy(m_hilDisabled); m_hilDisabled = nullptr; }

	// Create new ones
	m_hilDefault  = _CreateImageListFromResource(hInst, bmp.idDef, bmp.cx, crMask);
	m_hilHot      = _CreateImageListFromResource(hInst, bmp.idHot, bmp.cx, crMask);
	m_hilDisabled = _CreateDisabledImageList(hInst, bmp.idDef, bmp.cx, crMask, settings.theme);

	// Append system bitmap icons for Properties, Cut, Copy, Paste, Folder Options
	_AppendSystemBitmapIcons(m_hilDefault, m_hilHot, m_hilDisabled, bmp.cx, settings.theme, fIE55);

	SendMessage(m_hWndToolbar, TB_SETIMAGELIST, 0, (LPARAM)m_hilDefault);
	SendMessage(m_hWndToolbar, TB_SETHOTIMAGELIST, 0, (LPARAM)m_hilHot);
	if (m_hilDisabled)
		SendMessage(m_hWndToolbar, TB_SETDISABLEDIMAGELIST, 0, (LPARAM)m_hilDisabled);

	SendMessage(m_hWndToolbar, TB_AUTOSIZE, 0, 0);
	InvalidateRect(m_hWndToolbar, nullptr, TRUE);
}

// ================================================================================================
// ReloadSettings: Re-read settings from registry and update the toolbar live.
// Called when the user changes settings via the Customize Toolbar dialog.
//

void CStandardToolbar::ReloadSettings()
{
	if (!m_hWndToolbar)
		return;

	CEUtil::CESettings settings = CEUtil::GetCESettings();

	// Detect theme switch — reset toolbar layout and settings to new skin's defaults
	if (m_lastTheme != CLASSIC_EXPLORER_NONE && settings.theme != m_lastTheme)
	{
		// Delete saved layout so InitToolbarButtons loads the new skin's defaults
		DeleteSavedToolbarLayout();

		// Write the new skin's default icon size and text mode to registry
		TB_SKIN_DEFAULTS skinDef = GetDefaultSettings(settings.theme);
		CEUtil::CESettings newSettings(
			settings.theme,
			settings.showGoButton,
			settings.showAddressLabel,
			settings.showFullAddress,
			skinDef.smallIcons,
			skinDef.textLabelMode);
		CEUtil::WriteCESettings(newSettings);

		// Re-read settings with the newly written defaults
		settings = CEUtil::GetCESettings();

		// Remove all existing buttons
		int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
		for (int i = nButtons - 1; i >= 0; i--)
			SendMessage(m_hWndToolbar, TB_DELETEBUTTON, i, 0);

		// Rebuild image lists for the new theme
		m_lastTheme = settings.theme;
		RebuildImageLists();

		// Re-populate with new skin's default layout
		PreloadAllStrings();
		InitToolbarButtons();
	}
	else
	{
		// Normal settings change (no theme switch) — just rebuild image lists
		RebuildImageLists();
	}

	// Apply text label mode
	DWORD dwTextMode = (settings.textLabelMode <= 2) ? settings.textLabelMode : CE_TEXTMODE_SELECTIVE;
	ApplyTextLabelMode(dwTextMode);

	// Apply IE 5.5 style: toggle BTNS_SHOWTEXT on the History button (2K only)
	if (settings.theme == CLASSIC_EXPLORER_2K)
	{
		int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
		for (int i = 0; i < nButtons; i++)
		{
			TBBUTTON tbb = {};
			SendMessage(m_hWndToolbar, TB_GETBUTTON, i, (LPARAM)&tbb);
			if (tbb.idCommand == TBIDM_HISTORY)
			{
				TBBUTTONINFO tbi = { sizeof(tbi) };
				tbi.dwMask = TBIF_STYLE;
				tbi.fsStyle = tbb.fsStyle;
				if (settings.ie55Style == 1)
					tbi.fsStyle |= BTNS_SHOWTEXT;
				else
					tbi.fsStyle &= ~BTNS_SHOWTEXT;
				SendMessage(m_hWndToolbar, TB_SETBUTTONINFO, TBIDM_HISTORY, (LPARAM)&tbi);
				break;
			}
		}
	}

	// Notify the rebar parent to recalculate band sizes
	HWND hRebar = GetParent(m_hWndToolbar);
	if (hRebar && IsWindow(hRebar))
	{
		// Find our band index and update its size
		int nBands = (int)SendMessage(hRebar, RB_GETBANDCOUNT, 0, 0);
		for (int i = 0; i < nBands; i++)
		{
			REBARBANDINFO rbi = { sizeof(rbi) };
			rbi.fMask = RBBIM_CHILD;
			SendMessage(hRebar, RB_GETBANDINFO, i, (LPARAM)&rbi);
			if (rbi.hwndChild == m_hWndToolbar)
			{
				// Force rebar to re-query band size
				DESKBANDINFO dbi = {};
				dbi.dwMask = DBIM_MINSIZE | DBIM_ACTUAL;
				GetBandInfo(0, 0, &dbi);

				rbi.fMask = RBBIM_CHILDSIZE;
				rbi.cxMinChild = dbi.ptMinSize.x;
				rbi.cyMinChild = dbi.ptMinSize.y;
				SendMessage(hRebar, RB_SETBANDINFO, i, (LPARAM)&rbi);
				break;
			}
		}
	}
}
