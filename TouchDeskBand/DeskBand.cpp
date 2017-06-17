#include <windows.h>
#include <Wincodec.h>
#pragma comment(lib, "windowscodecs.lib")

#include <uxtheme.h>
#pragma comment(lib, "uxtheme.lib")

#include "DeskBand.h"
#include "resource.h"

#include <tchar.h>
#include <stdio.h>

//#include <winsock2.h>
//#include <ws2tcpip.h>
#pragma comment(lib, "IPHLPAPI.lib")

#include <Iphlpapi.h>

#include <time.h>    

//#include <iostream>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Audiopolicy.h>
#include <comdef.h>
#include <comip.h>

#define SAFE_RELEASE(punk)  \
              if ((punk) != NULL)  \
                { (punk)->Release(); (punk) = NULL; }

//#define CHECK_HR(hr)   \
//    if(FAILED(hr)) {    \
//        std::cout << "error" << std::endl;   \
//        return 0;   \
//                }
_COM_SMARTPTR_TYPEDEF(IMMDevice, __uuidof(IMMDevice));
_COM_SMARTPTR_TYPEDEF(IMMDeviceEnumerator, __uuidof(IMMDeviceEnumerator));
_COM_SMARTPTR_TYPEDEF(IAudioSessionManager2, __uuidof(IAudioSessionManager2));
_COM_SMARTPTR_TYPEDEF(IAudioSessionManager2, __uuidof(IAudioSessionManager2));
_COM_SMARTPTR_TYPEDEF(IAudioSessionEnumerator, __uuidof(IAudioSessionEnumerator));
_COM_SMARTPTR_TYPEDEF(IAudioSessionControl2, __uuidof(IAudioSessionControl2));
_COM_SMARTPTR_TYPEDEF(IAudioSessionControl, __uuidof(IAudioSessionControl));
_COM_SMARTPTR_TYPEDEF(ISimpleAudioVolume, __uuidof(ISimpleAudioVolume));
_COM_SMARTPTR_TYPEDEF(IAudioMeterInformation, __uuidof(IAudioMeterInformation));

IAudioSessionManager2Ptr CreateSessionManager()
{
	HRESULT hr = S_OK;

	static IMMDevicePtr pDevice;
	static IMMDeviceEnumeratorPtr pEnumerator;
	static IAudioSessionManager2Ptr pSessionManager;
	if (pSessionManager) {
		return pSessionManager;
	}
	else {
		SAFE_RELEASE(pDevice);
		SAFE_RELEASE(pEnumerator);
	}

	// Create the device enumerator.
	if (SUCCEEDED(hr = CoCreateInstance(__uuidof(MMDeviceEnumerator),
		NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator))) {
		if (SUCCEEDED(hr = pEnumerator->GetDefaultAudioEndpoint(
			eRender, eConsole, &pDevice))) {
			
			if (SUCCEEDED(hr = pDevice->Activate(
				__uuidof(IAudioSessionManager2), CLSCTX_ALL,
				NULL, (void**)&pSessionManager))) {
				return pSessionManager;
			}
			SAFE_RELEASE(pDevice);
		}
		SAFE_RELEASE(pEnumerator);
	}
	return NULL;
}


int getCurrentTime()
{
	SYSTEMTIME utc_time = { 0 };
	GetSystemTime(&utc_time);
	return utc_time.wMilliseconds + 1000 * utc_time.wSecond;
}


#define RECTWIDTH(x)   ((x).right - (x).left)
#define RECTHEIGHT(x)  ((x).bottom - (x).top)

extern ULONG        g_cDllRef;
extern HINSTANCE    g_hInst;
extern bool         g_canUnload;


extern CLSID CLSID_Touch;
static const WCHAR g_szDeskBandSampleClass[] = L"DeskBandSampleClass";

CDeskBand::CDeskBand() :
	m_cRef(1), m_pSite(NULL), m_fHasFocus(FALSE), m_fIsDirty(FALSE), m_dwBandID(0), m_hwnd(NULL), m_hwndParent(NULL)
{
}

CDeskBand::~CDeskBand()
{
	if (m_pSite)
	{
		m_pSite->Release();
	}
}

//
// IUnknown
//
STDMETHODIMP CDeskBand::QueryInterface(REFIID riid, void **ppv)
{
	HRESULT hr = S_OK;

	if (IsEqualIID(IID_IUnknown, riid) ||
		IsEqualIID(IID_IOleWindow, riid) ||
		IsEqualIID(IID_IDockingWindow, riid) ||
		IsEqualIID(IID_IDeskBand, riid) ||
		IsEqualIID(IID_IDeskBand2, riid))
	{
		*ppv = static_cast<IOleWindow *>(this);
	}
	else if (IsEqualIID(IID_IPersist, riid) ||
		IsEqualIID(IID_IPersistStream, riid))
	{
		*ppv = static_cast<IPersist *>(this);
	}
	else if (IsEqualIID(IID_IObjectWithSite, riid))
	{
		*ppv = static_cast<IObjectWithSite *>(this);
	}
	else if (IsEqualIID(IID_IInputObject, riid))
	{
		*ppv = static_cast<IInputObject *>(this);
	}
	else
	{
		hr = E_NOINTERFACE;
		*ppv = NULL;
	}

	if (*ppv)
	{
		AddRef();
	}

	return hr;
}

STDMETHODIMP_(ULONG) CDeskBand::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG) CDeskBand::Release()
{
	ULONG cRef = InterlockedDecrement(&m_cRef);
	if (0 == cRef && g_canUnload)
	{
		delete this;
	}
	return cRef;
}

//
// IOleWindow
//
STDMETHODIMP CDeskBand::GetWindow(HWND *phwnd)
{
	*phwnd = m_hwnd;
	return S_OK;
}

STDMETHODIMP CDeskBand::ContextSensitiveHelp(BOOL)
{
	return E_NOTIMPL;
}

//
// IDockingWindow
//
STDMETHODIMP CDeskBand::ShowDW(BOOL fShow)
{
	if (m_hwnd)
	{
		ShowWindow(m_hwnd, fShow ? SW_SHOW : SW_HIDE);
	}

	return S_OK;
}

STDMETHODIMP CDeskBand::CloseDW(DWORD)
{
	if (m_hwnd)
	{
		ShowWindow(m_hwnd, SW_HIDE);
		DestroyWindow(m_hwnd);
		m_hwnd = NULL;
	}

	return S_OK;
}

STDMETHODIMP CDeskBand::ResizeBorderDW(const RECT *, IUnknown *, BOOL)
{
	return E_NOTIMPL;
}

//
// IDeskBand
//
STDMETHODIMP CDeskBand::GetBandInfo(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO *pdbi)
{
	HRESULT hr = E_INVALIDARG;

	if (pdbi)
	{
		m_dwBandID = dwBandID;

		POINT pt;

		pt.x = dwViewMode == DBIF_VIEWMODE_NORMAL ? 70 : 35;
		pt.y = dwViewMode == DBIF_VIEWMODE_NORMAL ? 35 : 70;

		if (pdbi->dwMask & DBIM_MINSIZE)
		{
			pdbi->ptMinSize.x = pt.x;
			pdbi->ptMinSize.y = pt.y;
		}

		if (pdbi->dwMask & DBIM_MAXSIZE)
		{
			pdbi->ptMaxSize.y = -1;
		}

		if (pdbi->dwMask & DBIM_INTEGRAL)
		{
			pdbi->ptIntegral.y = 1;
		}

		if (pdbi->dwMask & DBIM_ACTUAL)
		{
			pdbi->ptActual.x = pt.x;
			pdbi->ptActual.y = pt.y;
		}

		if (pdbi->dwMask & DBIM_TITLE)
		{
			// Don't show title by removing this flag.
			pdbi->dwMask &= ~DBIM_TITLE;
		}

		if (pdbi->dwMask & DBIM_MODEFLAGS)
		{
			pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_VARIABLEHEIGHT;
		}

		if (pdbi->dwMask & DBIM_BKCOLOR)
		{
			// Use the default background color by removing this flag.
			pdbi->dwMask &= ~DBIM_BKCOLOR;
		}

		hr = S_OK;
	}

	return hr;
}

//
// IDeskBand2
//
STDMETHODIMP CDeskBand::CanRenderComposited(BOOL *pfCanRenderComposited)
{
	*pfCanRenderComposited = TRUE;

	return S_OK;
}

STDMETHODIMP CDeskBand::SetCompositionState(BOOL fCompositionEnabled)
{
	m_fCompositionEnabled = fCompositionEnabled;

	InvalidateRect(m_hwnd, NULL, TRUE);
	UpdateWindow(m_hwnd);

	return S_OK;
}

STDMETHODIMP CDeskBand::GetCompositionState(BOOL *pfCompositionEnabled)
{
	*pfCompositionEnabled = m_fCompositionEnabled;

	return S_OK;
}

//
// IPersist
//
STDMETHODIMP CDeskBand::GetClassID(CLSID *pclsid)
{
	*pclsid = CLSID_Touch;
	return S_OK;
}

//
// IPersistStream
//
STDMETHODIMP CDeskBand::IsDirty()
{
	return m_fIsDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CDeskBand::Load(IStream * /*pStm*/)
{
	return S_OK;
}

STDMETHODIMP CDeskBand::Save(IStream * /*pStm*/, BOOL fClearDirty)
{
	if (fClearDirty)
	{
		m_fIsDirty = FALSE;
	}

	return S_OK;
}

STDMETHODIMP CDeskBand::GetSizeMax(ULARGE_INTEGER * /*pcbSize*/)
{
	return E_NOTIMPL;
}

//
// IObjectWithSite
//
STDMETHODIMP CDeskBand::SetSite(IUnknown *pUnkSite)
{
	HRESULT hr = S_OK;

	m_hwndParent = NULL;

	if (m_pSite)
	{
		m_pSite->Release();
	}

	if (pUnkSite)
	{
		IOleWindow *pow;
		hr = pUnkSite->QueryInterface(IID_IOleWindow, reinterpret_cast<void **>(&pow));
		if (SUCCEEDED(hr))
		{
			hr = pow->GetWindow(&m_hwndParent);
			if (SUCCEEDED(hr))
			{
				WNDCLASSW wc = { 0 };
				wc.style = CS_HREDRAW | CS_VREDRAW;
				wc.hCursor = LoadCursor(NULL, IDC_ARROW);
				wc.hInstance = g_hInst;
				wc.lpfnWndProc = WndProc;
				wc.lpszClassName = g_szDeskBandSampleClass;
				wc.hbrBackground = CreateSolidBrush(RGB(255, 255, 0));

				RegisterClassW(&wc);

				CreateWindowExW(0,
					g_szDeskBandSampleClass,
					NULL,
					WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
					0,
					0,
					0,
					0,
					m_hwndParent,
					NULL,
					g_hInst,
					this);

				if (!m_hwnd)
				{
					hr = E_FAIL;
				}
			}

			pow->Release();
		}

		hr = pUnkSite->QueryInterface(IID_IInputObjectSite, reinterpret_cast<void **>(&m_pSite));
	}

	return hr;
}

STDMETHODIMP CDeskBand::GetSite(REFIID riid, void **ppv)
{
	HRESULT hr = E_FAIL;

	if (m_pSite)
	{
		hr = m_pSite->QueryInterface(riid, ppv);
	}
	else
	{
		*ppv = NULL;
	}

	return hr;
}

//
// IInputObject
//
STDMETHODIMP CDeskBand::UIActivateIO(BOOL fActivate, MSG *)
{
	if (fActivate)
	{
		SetFocus(m_hwnd);
	}

	return S_OK;
}

STDMETHODIMP CDeskBand::HasFocusIO()
{
	return m_fHasFocus ? S_OK : S_FALSE;
}

STDMETHODIMP CDeskBand::TranslateAcceleratorIO(MSG *)
{
	return S_FALSE;
};

void CDeskBand::OnFocus(const BOOL fFocus)
{
	m_fHasFocus = fFocus;

	if (m_pSite)
	{
		m_pSite->OnFocusChangeIS(static_cast<IOleWindow*>(this), m_fHasFocus);
	}
}









HHOOK mhook = 0;

#define MOUSEEVENTF_FROMTOUCH_NOPEN 0xFF515780 
#define MOUSEEVENTF_FROMTOUCH 0xFF515700
LRESULT CALLBACK mouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{

	static ULONG oinfo = 0;
	bool block = false;
	if (nCode == HC_ACTION) {
		MSLLHOOKSTRUCT* info = (MSLLHOOKSTRUCT *)lParam;

		HWND hwnd = GetForegroundWindow();
		TCHAR className[15];
		GetClassName(hwnd, className, 15);

		if (oinfo != info->dwExtraInfo) {
			oinfo = (ULONG)info->dwExtraInfo;
			if (wParam == WM_LBUTTONDOWN && info->dwExtraInfo == 0)
				block = true;
		}

		//Shell_TrayWnd
		if ((info->dwExtraInfo & MOUSEEVENTF_FROMTOUCH_NOPEN) == MOUSEEVENTF_FROMTOUCH_NOPEN && _tcscmp(className, L"Shell_TrayWnd"))
			block = true;

		if (block)
			return 1;
	}
	return CallNextHookEx(mhook, nCode, wParam, lParam);
}

IAudioMeterInformationPtr GetMeter(HWND hwnd) {
	if (!hwnd)
		return NULL;
	DWORD processId;
	GetWindowThreadProcessId(hwnd, &processId);
	if(!processId)
		return NULL;

	static IAudioSessionManager2Ptr mgr;
	static IAudioSessionEnumeratorPtr enumerator;
	static IAudioSessionControlPtr control;
	static IAudioSessionControl2Ptr control2;
	if (!mgr)
		mgr = CreateSessionManager();

	if (!mgr) {
		return NULL;
	}
	else {
		SAFE_RELEASE(control2);
		SAFE_RELEASE(control);
		SAFE_RELEASE(enumerator);
		//SAFE_RELEASE(mgr);
	}


	if (SUCCEEDED(mgr->GetSessionEnumerator(&enumerator))) {
		int sessionCount;
		if (SUCCEEDED(enumerator->GetCount(&sessionCount))) {
			for (int i = 0; i < sessionCount; i++) {
				if (SUCCEEDED(enumerator->GetSession(i, &control))) {
					
					if (SUCCEEDED(control->QueryInterface(__uuidof(IAudioSessionControl2), (void**)&control2))) {
						DWORD foundProcessId;
						if (SUCCEEDED(control2->GetProcessId(&foundProcessId))) {
							if (foundProcessId == processId) {
								IAudioMeterInformationPtr mtr;
								if (SUCCEEDED(control2->QueryInterface(_uuidof(IAudioMeterInformation), (void**)&mtr))) {
									return mtr;
								}
							}
						}
						SAFE_RELEASE(control2);
					}
					SAFE_RELEASE(control);
				}
			}
		}
		SAFE_RELEASE(enumerator);
	}
	return NULL;
}


bool showLyrics = false;
DWORD WINAPI NetEasyThread(LPVOID lParam) {
	HWND NetEasy = NULL;
	IAudioMeterInformationPtr mtr;
	while (IsWindow((HWND)lParam)) {
		//加载句柄
		if (!NetEasy || !IsWindow(NetEasy)) 
		{
			NetEasy = ::FindWindow(L"DesktopLyrics", NULL);
			SAFE_RELEASE(mtr);
		}
		else
		{
			if (!mtr) {
				mtr = GetMeter(NetEasy);
			}
			//if (!mtr) {
			//	Sleep(100000);
			//}

			if (mtr) {
				float v;
				if (SUCCEEDED(mtr->GetPeakValue(&v))) {
					showLyrics = v == 0 ? false : true;
				}
			}
			if (showLyrics) {
				RECT rgn;
				GetWindowRect(NetEasy, &rgn);
				int ht = rgn.bottom - rgn.top;
				SystemParametersInfo(SPI_GETWORKAREA, 0, (PVOID)&rgn, 0);
				SetWindowPos(NetEasy, 0, rgn.left, rgn.bottom - ht, rgn.right - rgn.left, ht, SWP_NOREDRAW | SWP_SHOWWINDOW);
			}
			else {
				ShowWindow(NetEasy, HIDE_WINDOW);
			}
			
		}
		Sleep(1000);
	}
	return 0;
}


void StartHook() {
	mhook = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, NULL, 0);
}

void EndHook() {
	UnhookWindowsHookEx(mhook);
	mhook = 0;
}


//bool hooking = false;
bool hover = false;
DWORD WINAPI hookThread(LPVOID lParam) {


	while (IsWindow((HWND)lParam)) {
		//if (hooking) {
		//	if(!mhook)
		//		mhook = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, NULL, 0);
		//}
		//else {
		//	if(mhook)
		//		UnhookWindowsHookEx(mhook);
		//	mhook = 0;
		//}

		POINT pt;
		GetCursorPos(&pt);
		RECT rect;
		GetWindowRect((HWND)lParam, &rect);
		bool h = PtInRect(&rect, pt);
		if(hover!=h)
			InvalidateRect((HWND)lParam, NULL, false);
		hover = h;

		Sleep(100);
	}
	return 0;
}









bool ImageFromIDResource(UINT nID, LPCTSTR sTR, Image *&pImg)
{
	HINSTANCE hInst = g_hInst;
	HRSRC hRsrc = ::FindResource(hInst, MAKEINTRESOURCE(nID), sTR); // type
	if (!hRsrc)
		return FALSE;
	// load resource into memory
	DWORD len = SizeofResource(hInst, hRsrc);
	BYTE* lpRsrc = (BYTE*)LoadResource(hInst, hRsrc);
	if (!lpRsrc)
		return FALSE;
	// Allocate global memory on which to create stream
	HGLOBAL m_hMem = GlobalAlloc(GMEM_FIXED, len);
	BYTE* pmem = (BYTE*)GlobalLock(m_hMem);
	memcpy(pmem, lpRsrc, len);
	GlobalUnlock(m_hMem);
	IStream* pstm;
	CreateStreamOnHGlobal(m_hMem, FALSE, &pstm);
	// load from stream
	pImg = Gdiplus::Image::FromStream(pstm);
	// free/release stuff
	pstm->Release();
	FreeResource(lpRsrc);
	GlobalFree(m_hMem);
	return TRUE;
}

void DrawString(Graphics& g, PWCHAR string, RectF rect, int size, StringAlignment sta)
{
	LOGFONTW lf;
	SystemParametersInfoW(SPI_GETICONTITLELOGFONT, sizeof(LOGFONTW), &lf, 0);
	SolidBrush  solidBrush(Color::White);
	FontFamily  fontFamily(lf.lfFaceName);
	Font        font(&fontFamily, (REAL)size, FontStyleRegular, UnitPixel);

	StringFormat sf(StringFormatFlags::StringFormatFlagsNoWrap);

	sf.SetAlignment(sta);
	sf.SetLineAlignment(StringAlignmentCenter);

	g.DrawString((WCHAR*)string, (INT)wcslen(string), &font, rect, &sf, &solidBrush);
}


void CDeskBand::OnPaint(const HDC hdcIn)
{
	if (g_canUnload)
		return;

	HDC hdc = hdcIn;
	PAINTSTRUCT ps;
	//static WCHAR szContent[] = L"DeskBand Sample";
	//static WCHAR szContentGlass[] = L"DeskBand Sample (Glass)";

	if (!hdc)
	{
		hdc = BeginPaint(m_hwnd, &ps);
	}

	if (hdc)
	{
		RECT rc;
		GetClientRect(m_hwnd, &rc);

		SIZE size;

		if (m_fCompositionEnabled)
		{
			HTHEME hTheme = OpenThemeData(NULL, L"BUTTON");
			if (hTheme)
			{

				HDC hdcPaint = NULL;
				HPAINTBUFFER hBufferedPaint = BeginBufferedPaint(hdc, &rc, BPBF_TOPDOWNDIB, NULL, &hdcPaint);

				Rect rect = Rect(rc.left, rc.top, RECTWIDTH(rc), RECTHEIGHT(rc));

				Graphics g(hdcPaint);

				g.Clear(Color::MakeARGB(mhook ? 30 : 0, 255, 255, 255));
				g.FillRectangle(&SolidBrush(Color::MakeARGB(hover ? 20 : 0, 255, 255, 255)), Rect(0, 0, rect.Width, rect.Height));

				if (!pimage)
					ImageFromIDResource(IDB_PNG1, _T("PNG"), pimage);

				if (pimage)
					g.DrawImage(pimage, (rect.Width - 32) / 2, (rect.Height - 32) / 2, 32, 32);

				EndBufferedPaint(hBufferedPaint, TRUE);

				CloseThemeData(hTheme);
			}
		}
		else
		{
			//SetBkColor(hdc, RGB(255, 255, 0));
			//GetTextExtentPointW(hdc, szContent, ARRAYSIZE(szContent), &size);
			//TextOutW(hdc,
			//	(RECTWIDTH(rc) - size.cx) / 2,
			//	(RECTHEIGHT(rc) - size.cy) / 2,
			//	szContent,
			//	ARRAYSIZE(szContent));
		}
	}

	if (!hdcIn)
	{
		EndPaint(m_hwnd, &ps);
	}
}

LRESULT CALLBACK CDeskBand::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = 0;

	CDeskBand *pDeskBand = reinterpret_cast<CDeskBand *>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

	switch (uMsg)
	{
	case WM_CREATE:
		g_canUnload = false;
		{
			static ULONG_PTR gdiplusToken;
			if (!gdiplusToken) {
				GdiplusStartupInput gdiplusStartupInput;
				GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
			}

			//创建NetEasy线程
			HANDLE hThread = CreateThread(NULL, 0, NetEasyThread, hwnd, 0, NULL);
			CloseHandle(hThread);
			//创建hook线程
			hThread = CreateThread(NULL, 0, hookThread, hwnd, 0, NULL);
			CloseHandle(hThread);
		}
		pDeskBand = reinterpret_cast<CDeskBand *>(reinterpret_cast<CREATESTRUCT *>(lParam)->lpCreateParams);
		pDeskBand->m_hwnd = hwnd;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pDeskBand));
		break;
	case WM_DESTROY:
		EndHook();
		g_canUnload = true;
		//GdiplusShutdown(gdiplusToken);
		break;
	case WM_LBUTTONUP:
		if (mhook) {
			EndHook();
		}
		else {
			StartHook();
		}
		break;
	case WM_SETFOCUS:
		pDeskBand->OnFocus(TRUE);
		break;

	case WM_KILLFOCUS:
		pDeskBand->OnFocus(FALSE);
		break;

	case WM_PAINT:
		pDeskBand->OnPaint(NULL);
		break;

	case WM_PRINTCLIENT:
		pDeskBand->OnPaint(reinterpret_cast<HDC>(wParam));
		break;

	case WM_ERASEBKGND:
		if (pDeskBand->m_fCompositionEnabled)
		{
			lResult = 1;
		}
		break;
	}

	if (uMsg != WM_ERASEBKGND)
	{
		lResult = DefWindowProc(hwnd, uMsg, wParam, lParam);
	}

	return lResult;
}
