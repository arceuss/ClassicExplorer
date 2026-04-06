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
#include "../AddressBar/AddressBar.h"

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

// (Settings change notification and customize dialog removed for NT4 conversion)

// ================================================================================================
// Top-level window subclass — no longer needed for NT4 (no settings change broadcasts)
// Kept as stub for potential future use.
//

static LRESULT CALLBACK TopLevelSubclassProc(
	HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
	UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
	if (uMsg == WM_NCDESTROY)
		RemoveWindowSubclass(hWnd, TopLevelSubclassProc, uIdSubclass);

	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

// ================================================================================================
// Container window ("band window") that hosts the combobox + toolbar side-by-side.
// The rebar owns this window; it parents both the ComboBoxEx and the toolbar control.
//

static const WCHAR g_szBandClass[] = L"ClassicExplorer.NT4Band";
static bool g_bBandClassRegistered = false;

static LRESULT CALLBACK BandContainerProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CStandardToolbar* pThis = (CStandardToolbar*)GetWindowLongPtr(hWnd, GWLP_USERDATA);

	switch (uMsg)
	{
	case WM_SIZE:
		if (pThis)
			pThis->LayoutBandChildren();
		return 0;

	case WM_COMMAND:
		// Forward WM_COMMAND from the combobox to the address bar message map
		if (pThis && pThis->GetAddressBar().IsWindow())
		{
			BOOL bHandled = FALSE;
			HWND hFrom = (HWND)lParam;
			// Forward toolbar commands to HandleCommand
			if (hFrom == pThis->GetToolbarHwnd())
			{
				pThis->HandleCommand(LOWORD(wParam));
				return 0;
			}
		}
		break;

	case WM_NOTIFY:
	{
		NMHDR* pnmh = (NMHDR*)lParam;
		if (pThis && pnmh)
		{
			// Forward toolbar notifications
			if (pnmh->hwndFrom == pThis->GetToolbarHwnd())
			{
				if (pnmh->code == TBN_DROPDOWN)
				{
					NMTOOLBAR* pnmtb = (NMTOOLBAR*)lParam;
					int idCmd = pnmtb->iItem;
					if (idCmd == TBIDM_BACK || idCmd == TBIDM_FORWARD)
					{
						pThis->ShowBackForwardMenu(idCmd, pnmtb);
						return TBDDRET_DEFAULT;
					}
				}
			}
			// Forward combobox notifications to the address bar
			CAddressBar& ab = pThis->GetAddressBar();
			if (ab.IsWindow())
			{
				BOOL bHandled = FALSE;
				LRESULT lr = ab.SendMessage(WM_NOTIFY, wParam, lParam);
			}
		}
		break;
	}

	case WM_ERASEBKGND:
		// Let the rebar draw the background through us
		return DefWindowProc(hWnd, uMsg, wParam, lParam);

	case WM_NCDESTROY:
		SetWindowLongPtr(hWnd, GWLP_USERDATA, 0);
		break;
	}

	return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

static void EnsureBandClassRegistered(HINSTANCE hInst)
{
	if (g_bBandClassRegistered)
		return;

	WNDCLASSEX wc = { sizeof(WNDCLASSEX) };
	wc.lpfnWndProc = BandContainerProc;
	wc.hInstance = hInst;
	wc.lpszClassName = g_szBandClass;
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
	RegisterClassEx(&wc);
	g_bBandClassRegistered = true;
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

	pDbi->dwModeFlags = DBIMF_USECHEVRON;

	// Compute combobox width the NT4 way: LOGPIXELSY * 2 (2 logical inches)
	if (m_comboWidth == 0)
	{
		HDC hdc = GetDC(nullptr);
		m_comboWidth = GetDeviceCaps(hdc, LOGPIXELSY) * 2;
		ReleaseDC(nullptr, hdc);
	}

	// Get toolbar button dimensions
	int tbHeight = 22;
	int tbIdealWidth = 0;
	if (m_hWndToolbar)
	{
		int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
		if (nButtons > 0)
		{
			RECT rcLast = {};
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, nButtons - 1, (LPARAM)&rcLast);
			tbIdealWidth = rcLast.right;

			RECT rcFirst = {};
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, 0, (LPARAM)&rcFirst);
			tbHeight = rcFirst.bottom - rcFirst.top;
		}
		else
		{
			LONG lSize = (LONG)SendMessage(m_hWndToolbar, TB_GETBUTTONSIZE, 0, 0);
			tbHeight = HIWORD(lSize);
		}
	}

	// NT4 non-flat toolbar: iYPos=2 top + g_cxEdge=2 bottom = 4px total
	int bandHeight = max(tbHeight + 4, 22);

	if (pDbi->dwMask & DBIM_MINSIZE)
	{
		// Minimum: combobox + separator + one button
		pDbi->ptMinSize.x = m_comboWidth + 4 + 24;
		pDbi->ptMinSize.y = bandHeight;
	}

	if (pDbi->dwMask & DBIM_MAXSIZE)
	{
		pDbi->ptMaxSize.x = 0;
		pDbi->ptMaxSize.y = bandHeight;
	}

	if (pDbi->dwMask & DBIM_INTEGRAL)
	{
		pDbi->ptIntegral.x = 0;
		pDbi->ptIntegral.y = 1;
	}

	if (pDbi->dwMask & DBIM_ACTUAL)
	{
		// Request full width: fixed combobox + gap + all buttons
		pDbi->ptActual.x = m_comboWidth + 4 + tbIdealWidth;
		pDbi->ptActual.y = bandHeight;
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

	*hWnd = m_hWndBand;
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
	if (m_hWndBand)
		ShowWindow(m_hWndBand, fShow ? SW_SHOW : SW_HIDE);
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

		// Pass browser interfaces to the embedded address bar
		m_addressBar.SetBrowsers(m_pShellBrowser, m_pWebBrowser);
		m_addressBar.InitComboBox();

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

	// Update the embedded address bar
	m_addressBar.HandleNavigate();

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
	if (m_hWndBand && (hFocus == m_hWndBand || IsChild(m_hWndBand, hFocus)))
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

// Helper: Load a bitmap resource and create an image list.
// Loads without LR_CREATEDIBSECTION so RLE-compressed bitmaps decompress properly.

static HIMAGELIST _CreateImageListFromResource(HINSTANCE hInst, UINT idBitmap, int cx, COLORREF crMask)
{
	HBITMAP hbm = (HBITMAP)LoadImage(hInst, MAKEINTRESOURCE(idBitmap),
		IMAGE_BITMAP, 0, 0, 0);
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

	HIMAGELIST himl = ImageList_Create(cx, cy, ILC_COLOR4 | ILC_MASK, cInitial, 0);
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

// (Disabled image list generation removed for NT4 — comctl32 handles disabled
//  rendering natively using PSDPxax emboss for non-alpha bitmaps.)

HRESULT CStandardToolbar::CreateToolbarWindow(HWND hWndParent)
{
	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();

	// Register and create the container window that holds both controls
	EnsureBandClassRegistered(hInst);

	m_hWndBand = CreateWindowEx(
		0,
		g_szBandClass,
		nullptr,
		WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
		0, 0, 0, 0,
		hWndParent,
		nullptr,
		hInst,
		nullptr);

	if (!m_hWndBand)
		return E_FAIL;

	SetWindowLongPtr(m_hWndBand, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(this));

	// Create the embedded address bar (ComboBoxEx) as a child of the container
	m_addressBar.Create(m_hWndBand, nullptr, nullptr, WS_CHILD);
	if (!m_addressBar.IsWindow())
		return E_FAIL;

	// NT4 toolbar: raised (non-flat), no rearranging, tooltips only (no text labels)
	DWORD dwTbStyle = WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS |
		TBSTYLE_TOOLTIPS | CCS_NODIVIDER | CCS_NOPARENTALIGN | CCS_NORESIZE;

	// Create the toolbar control as a child of the container
	m_hWndToolbar = CreateWindowEx(
		0,
		TOOLBARCLASSNAME,
		nullptr,
		dwTbStyle,
		0, 0, 0, 0,
		m_hWndBand,
		nullptr,
		hInst,
		nullptr);

	if (!m_hWndToolbar)
		return E_FAIL;

	// Required: set the struct size for the toolbar
	SendMessage(m_hWndToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

	// No text rows (NT4 toolbar uses tooltips only, no text labels)
	SendMessage(m_hWndToolbar, TB_SETMAXTEXTROWS, 0, 0);

	// Load NT4 comctl32 bitmap strips (STD + VIEW) into a single image list
	COLORREF crMask = RGB(192, 192, 192);  // NT4 bitmaps use button face gray as background

	// NT4 always uses small (16x16) toolbar icons
	bool fSmallIcons = true;
	TB_SKIN_BITMAPS bmp = GetActiveBitmapIDs(fSmallIcons);

	// Create image list from the STD strip (15 glyphs)
	m_hilDefault = _CreateImageListFromResource(hInst, bmp.idStd, bmp.cx, crMask);
	if (!m_hilDefault)
		return E_FAIL;

	// Append the VIEW strip (12 glyphs) to the same image list
	HBITMAP hbmView = (HBITMAP)LoadImage(hInst, MAKEINTRESOURCE(bmp.idView),
		IMAGE_BITMAP, 0, 0, 0);
	if (hbmView)
	{
		ImageList_AddMasked(m_hilDefault, hbmView, crMask);
		DeleteObject(hbmView);
	}

	SendMessage(m_hWndToolbar, TB_SETIMAGELIST, 0, (LPARAM)m_hilDefault);

	// Pre-load all button strings into the toolbar's string pool
	PreloadAllStrings();

	// Populate buttons (fixed NT4 layout)
	InitToolbarButtons();

	// Autosize the toolbar
	SendMessage(m_hWndToolbar, TB_AUTOSIZE, 0, 0);

	// Install a subclass for message handling
	SetWindowSubclass(m_hWndToolbar, ToolbarSubclassProc, 0, reinterpret_cast<DWORD_PTR>(this));

	// Show both child windows and perform initial layout
	ShowWindow(m_addressBar.m_hWnd, SW_SHOW);
	ShowWindow(m_hWndToolbar, SW_SHOW);
	LayoutBandChildren();

	// Install rebar subclass to intercept WM_COMMAND and WM_NOTIFY from our toolbar
	InstallRebarHook();

	return S_OK;
}

void CStandardToolbar::InitToolbarButtons()
{
	if (!m_hWndToolbar)
		return;

	// Build TBBUTTON array from fixed NT4 layout
	TBBUTTON tbb[32] = {};
	int nButtons = 0;

	for (int i = 0; i < c_nDefaultLayout_NT4 && nButtons < 32; i++)
	{
		if (c_defaultLayout_NT4[i] == 0)
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
			int idx = FindButtonCatalogIndex(c_defaultLayout_NT4[i]);
			if (idx < 0) continue;

			const TB_BUTTON_DEF& def = c_tbAllButtons_NT4[idx];
			tbb[nButtons].iBitmap = def.iBitmap;
			tbb[nButtons].idCommand = def.idCommand;
			tbb[nButtons].fsState = TBSTATE_ENABLED;
			tbb[nButtons].fsStyle = def.fsStyle;
			tbb[nButtons].dwData = 0;
			tbb[nButtons].iString = m_iStringPool[idx];
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
	if (m_addressBar.IsWindow())
		m_addressBar.DestroyWindow();
	if (m_hWndToolbar && IsWindow(m_hWndToolbar))
	{
		DestroyWindow(m_hWndToolbar);
		m_hWndToolbar = nullptr;
	}
	if (m_hWndBand && IsWindow(m_hWndBand))
	{
		DestroyWindow(m_hWndBand);
		m_hWndBand = nullptr;
	}
}

// ================================================================================================
// Pre-load all button strings into the toolbar's string pool.
// Must be called before InitToolbarButtons() so that string indices are available.
//

void CStandardToolbar::PreloadAllStrings()
{
	memset(m_iStringPool, -1, sizeof(m_iStringPool));

	HINSTANCE hInst = _AtlBaseModule.GetModuleInstance();
	for (int i = 0; i < c_nAllButtons_NT4; i++)
	{
		WCHAR szText[64] = {};
		LoadStringW(hInst, c_tbAllButtons_NT4[i].idsString, szText, ARRAYSIZE(szText));
		m_iStringPool[i] = (int)SendMessage(m_hWndToolbar, TB_ADDSTRING, 0, (LPARAM)szText);
	}
}

// (SaveToolbarLayout, LoadToolbarLayout, DeleteSavedToolbarLayout removed — NT4 uses fixed layout)

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

	// Sync view mode check buttons
	UpdateViewModeChecks();

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
				break;
			}

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
	case TBIDM_BACK:                OnBack(); break;
	case TBIDM_FORWARD:             OnForward(); break;
	case TBIDM_PREVIOUSFOLDER:      OnUpOneLevel(); break;
	case TBIDM_SEARCH:              OnSearch(); break;
	case TBIDM_ALLFOLDERS:          OnFolders(); break;
	case TBIDM_CONNECT:             OnMapNetDrive(); break;
	case TBIDM_DISCONNECT:          OnDisconnectNetDrive(); break;
	case TBIDM_CUT:                 OnCut(); break;
	case TBIDM_COPY:                OnCopy(); break;
	case TBIDM_PASTE:               OnPaste(); break;
	case TBIDM_UNDO:                OnUndo(); break;
	case TBIDM_DELETE:              OnDelete(); break;
	case TBIDM_PROPERTIES:          OnProperties(); break;
	case TBIDM_STOPDOWNLOAD:        OnStop(); break;
	case TBIDM_REFRESH:             OnRefresh(); break;
	case TBIDM_HOME:                OnHome(); break;
	case TBIDM_FAVORITES:           OnFavorites(); break;
	case TBIDM_HISTORY:             OnHistory(); break;
	case TBIDM_THEATER:             OnFullScreen(); break;
	// NT4 view mode buttons (BTNS_CHECKGROUP)
	case TBIDM_VIEW_LARGEICONS:     SetClassicViewMode(FVM_ICON); break;
	case TBIDM_VIEW_SMALLICONS:     SetClassicViewMode(FVM_SMALLICON); break;
	case TBIDM_VIEW_LIST:           SetClassicViewMode(FVM_LIST); break;
	case TBIDM_VIEW_DETAILS:        SetClassicViewMode(FVM_DETAILS); break;
	}
}

// Install the rebar subclass (called from CreateToolbarWindow after toolbar is created)
// With the container window architecture, the rebar is the grandparent of the toolbar.
// Most command routing is handled by BandContainerProc, but we still subclass the rebar
// for any notifications that bubble up.
void CStandardToolbar::InstallRebarHook()
{
	if (m_hWndBand)
	{
		HWND hRebar = GetParent(m_hWndBand);
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

// (OnViewsDropdown and CycleViewMode removed — NT4 uses individual view mode buttons)

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

// (ApplyTextLabelMode, RebuildImageLists, ReloadSettings removed — NT4 uses fixed layout and style)

// ================================================================================================
// LayoutBandChildren: Position the combobox and toolbar inside the container band.
// Layout: [ComboBoxEx | gap | Toolbar buttons]
//

void CStandardToolbar::LayoutBandChildren()
{
	if (!m_hWndBand)
		return;

	RECT rcBand = {};
	GetClientRect(m_hWndBand, &rcBand);
	int bandW = rcBand.right;
	int bandH = rcBand.bottom;

	if (bandW <= 0 || bandH <= 0)
		return;

	constexpr int GAP = 4;  // gap between combobox and toolbar

	// NT4 combobox width: LOGPIXELSY * 2 (2 logical inches, 192px @ 96 DPI)
	if (m_comboWidth == 0)
	{
		HDC hdc = GetDC(nullptr);
		m_comboWidth = GetDeviceCaps(hdc, LOGPIXELSY) * 2;
		ReleaseDC(nullptr, hdc);
	}
	int comboW = m_comboWidth;

	// Get actual toolbar width from button rects
	int tbWidth = 0;
	if (m_hWndToolbar)
	{
		int nButtons = (int)SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0);
		if (nButtons > 0)
		{
			RECT rcLast = {};
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, nButtons - 1, (LPARAM)&rcLast);
			tbWidth = rcLast.right;
		}
	}

	// Position the CAddressBar wrapper window (fixed width, left side)
	HWND hAddrWnd = m_addressBar.m_hWnd;
	if (hAddrWnd && IsWindow(hAddrWnd))
	{
		// Vertically center the combobox within the band (NT4: PositionDrivesCombo)
		RECT rcCombo = {};
		GetWindowRect(hAddrWnd, &rcCombo);
		int comboH = rcCombo.bottom - rcCombo.top;
		if (comboH < 22) comboH = 22;
		int comboY = (bandH - comboH) / 2;
		if (comboY < 0) comboY = 0;
		SetWindowPos(hAddrWnd, nullptr, 0, comboY, comboW, comboH,
			SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

		// Resize the ComboBoxEx to fill the wrapper (drop-down uses comboW for height like NT4)
		HWND hCombo = m_addressBar.GetToolbar();
		if (hCombo && IsWindow(hCombo))
		{
			SetWindowPos(hCombo, nullptr, 0, 0, comboW, comboW,
				SWP_NOZORDER | SWP_NOACTIVATE);
		}
	}

	// Position the toolbar to the right of the combobox
	if (m_hWndToolbar && IsWindow(m_hWndToolbar))
	{
		int tbX = comboW + GAP;
		// Get toolbar button height for vertical centering
		int tbHeight = bandH;
		RECT rcBtn = {};
		if (SendMessage(m_hWndToolbar, TB_BUTTONCOUNT, 0, 0))
		{
			SendMessage(m_hWndToolbar, TB_GETITEMRECT, 0, (LPARAM)&rcBtn);
			tbHeight = rcBtn.bottom - rcBtn.top;
		}
		int tbY = (bandH - tbHeight) / 2;
		if (tbY < 0) tbY = 0;
		SetWindowPos(m_hWndToolbar, nullptr, tbX, tbY, tbWidth, tbHeight,
			SWP_NOZORDER | SWP_NOACTIVATE);
	}
}

// ================================================================================================
// UpdateViewModeChecks: sync the BTNS_CHECKGROUP view mode buttons with the
// current folder's view mode. Called from UpdateButtonStates on each navigation.
//

void CStandardToolbar::UpdateViewModeChecks()
{
	if (!m_hWndToolbar || !m_pShellBrowser)
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

	int checkedCmd = 0;
	switch (fvm)
	{
	case FVM_ICON:
		// FVM_ICON with >=32px = Large Icons, <32px = treat as Small Icons
		checkedCmd = (iIconSize >= 32) ? TBIDM_VIEW_LARGEICONS : TBIDM_VIEW_SMALLICONS;
		break;
	case FVM_SMALLICON:
		// Shell converts 32px FVM_ICON to FVM_SMALLICON; disambiguate by size
		checkedCmd = (iIconSize >= 32) ? TBIDM_VIEW_LARGEICONS : TBIDM_VIEW_SMALLICONS;
		break;
	case FVM_LIST:
		checkedCmd = TBIDM_VIEW_LIST;
		break;
	case FVM_DETAILS:
		checkedCmd = TBIDM_VIEW_DETAILS;
		break;
	default:
		checkedCmd = TBIDM_VIEW_LARGEICONS;
		break;
	}

	SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_VIEW_LARGEICONS, MAKELONG(checkedCmd == TBIDM_VIEW_LARGEICONS, 0));
	SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_VIEW_SMALLICONS, MAKELONG(checkedCmd == TBIDM_VIEW_SMALLICONS, 0));
	SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_VIEW_LIST, MAKELONG(checkedCmd == TBIDM_VIEW_LIST, 0));
	SendMessage(m_hWndToolbar, TB_CHECKBUTTON, TBIDM_VIEW_DETAILS, MAKELONG(checkedCmd == TBIDM_VIEW_DETAILS, 0));
}
