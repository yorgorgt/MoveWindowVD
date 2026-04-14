// MoveWindowVD - Move active window to adjacent virtual desktop
// Windows 11  |  Ctrl+Shift+Win+Left / Ctrl+Shift+Win+Right
// Supports OS builds 22000 (21H2) and 26100 (24H2)

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <shellapi.h>
#include <objbase.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")

// ----------------------------------------------------------------
// OS build detection via RtlGetVersion
// ----------------------------------------------------------------
typedef LONG(WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);

static DWORD GetOSBuildNumber() {
    HMODULE hNtdll = GetModuleHandleW(L"ntdll.dll");
    if (!hNtdll) return 0;
    auto pRtlGetVersion = reinterpret_cast<RtlGetVersionPtr>(
        GetProcAddress(hNtdll, "RtlGetVersion"));
    if (!pRtlGetVersion) return 0;
    RTL_OSVERSIONINFOW vi = {};
    vi.dwOSVersionInfoSize = sizeof(vi);
    if (pRtlGetVersion(&vi) != 0) return 0;
    return vi.dwBuildNumber;
}

// ----------------------------------------------------------------
// CLSID / IID definitions for undocumented COM interfaces
// ----------------------------------------------------------------
static const CLSID CLSID_ImmersiveShell = {
    0xC2F03A33, 0x21F5, 0x47FA,
    {0xB4, 0xBB, 0x15, 0x63, 0x62, 0xA2, 0xF2, 0x39}
};
static const CLSID CLSID_VDMInternal = {
    0xC5E0CDCA, 0x7B6E, 0x41B2,
    {0x9F, 0xC4, 0xD9, 0x39, 0x75, 0xCC, 0x46, 0x7B}
};
static const IID IID_IServiceProvider10 = {
    0x6D5140C1, 0x7436, 0x11CE,
    {0x80, 0x34, 0x00, 0xAA, 0x00, 0x60, 0x09, 0xFA}
};
static const IID IID_IObjectArray = {
    0x92CA9DCD, 0x5622, 0x4BBA,
    {0xA8, 0x05, 0x5E, 0x9F, 0x54, 0x1B, 0xD8, 0xC9}
};
static const IID IID_IApplicationViewCollection = {
    0x1841C6D7, 0x4F9D, 0x42C0,
    {0xAF, 0x41, 0x87, 0x47, 0x53, 0x8F, 0x10, 0xE5}
};

// Build 22000 (21H2)
static const IID IID_IVirtualDesktop_22000 = {
    0x536D3495, 0xB208, 0x4CC9,
    {0xAE, 0x26, 0xDE, 0x81, 0x11, 0x27, 0x5B, 0xF8}
};
static const IID IID_IVDMInternal_22000 = {
    0xB2F925B9, 0x5A0F, 0x4D2E,
    {0x9F, 0x4D, 0x2B, 0x15, 0x07, 0x59, 0x3C, 0x10}
};

// Build 26100 (24H2)
static const IID IID_IVirtualDesktop_26100 = {
    0x3F07F4BE, 0xB107, 0x441A,
    {0xAF, 0x0F, 0x39, 0xD8, 0x25, 0x29, 0x07, 0x2C}
};
static const IID IID_IVDMInternal_26100 = {
    0x53F5CA0B, 0x158F, 0x4124,
    {0x90, 0x0C, 0x05, 0x71, 0x58, 0x06, 0x0B, 0x27}
};

// ----------------------------------------------------------------
// COM interface structures (vtable layout)
// ----------------------------------------------------------------
struct IServiceProvider10 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE QueryService(
        REFGUID guidService, REFIID riid, void** ppvObject) = 0;
};

struct IObjectArray : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetCount(UINT* pcObjects) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetAt(
        UINT uiIndex, REFIID riid, void** ppv) = 0;
};

struct IVirtualDesktop : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetID(GUID* pGuid) = 0;
};

struct IApplicationView : public IUnknown {};

struct IApplicationViewCollection : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetViews(IObjectArray** ppArray) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewsByZOrder(IObjectArray** ppArray) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewsByAppUserModelId(
        LPCWSTR id, IObjectArray** ppArray) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetViewForHwnd(
        HWND hwnd, IApplicationView** ppView) = 0;
};

// Build 22000: methods take HMONITOR parameter
struct IVirtualDesktopManagerInternal_22000 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetCount(
        HMONITOR hMon, UINT* pCount) = 0;
    virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(
        IApplicationView* pView, IVirtualDesktop* pDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(
        IApplicationView* pView, BOOL* pCanMove) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(
        HMONITOR hMon, IVirtualDesktop** ppDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDesktops(
        HMONITOR hMon, IObjectArray** ppArray) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(
        IVirtualDesktop* pDesktop, int uDirection,
        IVirtualDesktop** ppAdjacentDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(
        HMONITOR hMon, IVirtualDesktop* pDesktop) = 0;
};

// Build 26100: methods without HMONITOR
struct IVirtualDesktopManagerInternal_26100 : public IUnknown {
    virtual HRESULT STDMETHODCALLTYPE GetCount(UINT* pCount) = 0;
    virtual HRESULT STDMETHODCALLTYPE MoveViewToDesktop(
        IApplicationView* pView, IVirtualDesktop* pDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE CanViewMoveDesktops(
        IApplicationView* pView, BOOL* pCanMove) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCurrentDesktop(
        IVirtualDesktop** ppDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetDesktops(
        IObjectArray** ppArray) = 0;
    virtual HRESULT STDMETHODCALLTYPE GetAdjacentDesktop(
        IVirtualDesktop* pDesktop, UINT uDirection,
        IVirtualDesktop** ppAdjacentDesktop) = 0;
    virtual HRESULT STDMETHODCALLTYPE SwitchDesktop(
        IVirtualDesktop* pDesktop) = 0;
};

// ----------------------------------------------------------------
// VDMWrapper - abstracts vtable differences between OS builds
// ----------------------------------------------------------------
class VDMWrapper {
    IUnknown* m_pRaw = nullptr;
    bool m_is22000 = false;

    IVirtualDesktopManagerInternal_22000* as22000() {
        return static_cast<IVirtualDesktopManagerInternal_22000*>(m_pRaw);
    }
    IVirtualDesktopManagerInternal_26100* as26100() {
        return static_cast<IVirtualDesktopManagerInternal_26100*>(m_pRaw);
    }
public:
    void Init(IUnknown* pRaw, bool is22000) {
        m_pRaw = pRaw;
        m_is22000 = is22000;
    }
    HRESULT GetCurrentDesktop(IVirtualDesktop** pp) {
        return m_is22000 ? as22000()->GetCurrentDesktop(NULL, pp)
                         : as26100()->GetCurrentDesktop(pp);
    }
    HRESULT GetDesktops(IObjectArray** pp) {
        return m_is22000 ? as22000()->GetDesktops(NULL, pp)
                         : as26100()->GetDesktops(pp);
    }
    HRESULT MoveViewToDesktop(IApplicationView* pView, IVirtualDesktop* pDesktop) {
        return m_is22000 ? as22000()->MoveViewToDesktop(pView, pDesktop)
                         : as26100()->MoveViewToDesktop(pView, pDesktop);
    }
    HRESULT SwitchDesktop(IVirtualDesktop* pDesktop) {
        return m_is22000 ? as22000()->SwitchDesktop(NULL, pDesktop)
                         : as26100()->SwitchDesktop(pDesktop);
    }
    void Release() {
        if (m_pRaw) { m_pRaw->Release(); m_pRaw = nullptr; }
    }
};

// ----------------------------------------------------------------
// Constants
// ----------------------------------------------------------------
enum { HOTKEY_MOVE_LEFT = 1, HOTKEY_MOVE_RIGHT = 2 };
enum { WM_TRAYICON = WM_USER + 1 };
enum { IDM_EXIT = 1001 };

// ----------------------------------------------------------------
// Globals
// ----------------------------------------------------------------
static VDMWrapper                  g_vdm;
static IApplicationViewCollection* g_pViewCollection  = nullptr;
static bool                        g_is22000          = false;

// ----------------------------------------------------------------
// Move the foreground window to the adjacent desktop
// ----------------------------------------------------------------
static bool MoveActiveWindow(bool left) {
    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return false;

    IApplicationView* pView = nullptr;
    if (FAILED(g_pViewCollection->GetViewForHwnd(hwnd, &pView)))
        return false;

    IVirtualDesktop* pCurrent = nullptr;
    if (FAILED(g_vdm.GetCurrentDesktop(&pCurrent))) {
        pView->Release();
        return false;
    }

    IObjectArray* pDesktops = nullptr;
    if (FAILED(g_vdm.GetDesktops(&pDesktops))) {
        pCurrent->Release();
        pView->Release();
        return false;
    }

    // Select the correct IID_IVirtualDesktop for this OS build
    const IID& iidVD = g_is22000 ? IID_IVirtualDesktop_22000
                                 : IID_IVirtualDesktop_26100;

    // Find current desktop index by COM identity comparison
    UINT count = 0;
    pDesktops->GetCount(&count);

    IUnknown* pCurUnk = nullptr;
    pCurrent->QueryInterface(IID_IUnknown, reinterpret_cast<void**>(&pCurUnk));

    int currentIndex = -1;
    for (UINT i = 0; i < count; i++) {
        IVirtualDesktop* pDesk = nullptr;
        if (SUCCEEDED(pDesktops->GetAt(i, iidVD,
                reinterpret_cast<void**>(&pDesk)))) {
            IUnknown* pDeskUnk = nullptr;
            pDesk->QueryInterface(IID_IUnknown, reinterpret_cast<void**>(&pDeskUnk));
            bool match = (pCurUnk == pDeskUnk);
            pDeskUnk->Release();
            pDesk->Release();
            if (match) { currentIndex = static_cast<int>(i); break; }
        }
    }
    pCurUnk->Release();
    pCurrent->Release();

    if (currentIndex < 0) {
        pDesktops->Release();
        pView->Release();
        return false;
    }

    int targetIndex = left ? currentIndex - 1 : currentIndex + 1;
    if (targetIndex < 0 || targetIndex >= static_cast<int>(count)) {
        pDesktops->Release();
        pView->Release();
        return false;
    }

    IVirtualDesktop* pTarget = nullptr;
    HRESULT hr = pDesktops->GetAt(static_cast<UINT>(targetIndex),
        iidVD, reinterpret_cast<void**>(&pTarget));
    pDesktops->Release();

    if (FAILED(hr)) {
        pView->Release();
        return false;
    }

    hr = g_vdm.MoveViewToDesktop(pView, pTarget);
    if (SUCCEEDED(hr))
        g_vdm.SwitchDesktop(pTarget);

    pTarget->Release();
    pView->Release();
    return SUCCEEDED(hr);
}

// ----------------------------------------------------------------
// Window procedure
// ----------------------------------------------------------------
static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_CREATE:
        if (!RegisterHotKey(hwnd, HOTKEY_MOVE_LEFT,  MOD_CONTROL | MOD_SHIFT | MOD_WIN, VK_LEFT) ||
            !RegisterHotKey(hwnd, HOTKEY_MOVE_RIGHT, MOD_CONTROL | MOD_SHIFT | MOD_WIN, VK_RIGHT)) {
            MessageBoxW(nullptr,
                L"ホットキー (Ctrl+Shift+Win+Arrow) の登録に失敗しました。\n"
                L"同じキーを使うアプリが起動していないか確認してください。",
                L"MoveWindowVD", MB_ICONERROR);
            return -1;
        }
        return 0;

    case WM_HOTKEY:
        if (wp == HOTKEY_MOVE_LEFT || wp == HOTKEY_MOVE_RIGHT)
            MoveActiveWindow(wp == HOTKEY_MOVE_LEFT);
        return 0;

    case WM_TRAYICON:
        if (LOWORD(lp) == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            HMENU hMenu = CreatePopupMenu();
            AppendMenuW(hMenu, MF_STRING, IDM_EXIT, L"終了 (&X)");
            SetForegroundWindow(hwnd);
            TrackPopupMenu(hMenu, TPM_RIGHTALIGN, pt.x, pt.y, 0, hwnd, nullptr);
            DestroyMenu(hMenu);
        }
        return 0;

    case WM_COMMAND:
        if (LOWORD(wp) == IDM_EXIT)
            DestroyWindow(hwnd);
        return 0;

    case WM_DESTROY:
        UnregisterHotKey(hwnd, HOTKEY_MOVE_LEFT);
        UnregisterHotKey(hwnd, HOTKEY_MOVE_RIGHT);
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ----------------------------------------------------------------
// Entry point
// ----------------------------------------------------------------
int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int) {
    // Single-instance guard
    HANDLE hMutex = CreateMutexW(nullptr, TRUE, L"MoveWindowVD_Mutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBoxW(nullptr, L"MoveWindowVD は既に起動しています。",
            L"MoveWindowVD", MB_ICONINFORMATION);
        return 0;
    }

    // Detect OS build and select IIDs
    DWORD build = GetOSBuildNumber();
    g_is22000 = (build < 22631);
    const IID& iidVD  = g_is22000 ? IID_IVDMInternal_22000 : IID_IVDMInternal_26100;

    // COM - get undocumented interfaces via ImmersiveShell
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    IServiceProvider10* pShell = nullptr;
    if (FAILED(CoCreateInstance(CLSID_ImmersiveShell, nullptr, CLSCTX_LOCAL_SERVER,
            IID_IServiceProvider10, reinterpret_cast<void**>(&pShell)))) {
        MessageBoxW(nullptr,
            L"仮想デスクトップ COM の初期化に失敗しました。\n"
            L"Windows 10 以降で実行してください。",
            L"MoveWindowVD", MB_ICONERROR);
        CoUninitialize();
        return 1;
    }

    IUnknown* pManagerRaw = nullptr;
    if (FAILED(pShell->QueryService(CLSID_VDMInternal, iidVD,
            reinterpret_cast<void**>(&pManagerRaw))) ||
        FAILED(pShell->QueryService(IID_IApplicationViewCollection,
            IID_IApplicationViewCollection,
            reinterpret_cast<void**>(&g_pViewCollection)))) {
        MessageBoxW(nullptr,
            L"仮想デスクトップ内部 COM の取得に失敗しました。",
            L"MoveWindowVD", MB_ICONERROR);
        if (pManagerRaw) pManagerRaw->Release();
        pShell->Release();
        CoUninitialize();
        return 1;
    }
    g_vdm.Init(pManagerRaw, g_is22000);
    pShell->Release();

    // Window class & hidden message window
    WNDCLASSW wc = {};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance      = hInst;
    wc.lpszClassName  = L"MoveWindowVD";
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(0, L"MoveWindowVD", nullptr, 0,
        0, 0, 0, 0, HWND_MESSAGE, nullptr, hInst, nullptr);
    if (!hwnd) {
        g_pViewCollection->Release();
        g_vdm.Release();
        CoUninitialize();
        return 1;
    }

    // System tray icon
    NOTIFYICONDATAW nid = {};
    nid.cbSize          = sizeof(nid);
    nid.hWnd            = hwnd;
    nid.uID             = 1;
    nid.uFlags          = NIF_ICON | NIF_TIP | NIF_MESSAGE;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon           = LoadIconW(nullptr, IDI_APPLICATION);
    wcscpy_s(nid.szTip, L"MoveWindowVD (Ctrl+Shift+Win+Arrow)");
    Shell_NotifyIconW(NIM_ADD, &nid);

    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    // Cleanup
    Shell_NotifyIconW(NIM_DELETE, &nid);
    g_pViewCollection->Release();
    g_vdm.Release();
    CoUninitialize();
    if (hMutex) { ReleaseMutex(hMutex); CloseHandle(hMutex); }
    return 0;
}
