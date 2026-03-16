// main.cpp
#define UNICODE
#define _UNICODE
#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <string>
#include <map>
#include <vector>
#include <iostream>

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")

// Registry path/key for persistence
static const wchar_t* kRegKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\ThemeManager";
static const wchar_t* kRegValue = L"MsStylesPath";

// Globals
HWND g_hTree = nullptr;
std::wstring g_msstylesPath;

// Save msstyles path to registry
void SaveMsStylesPath(const std::wstring& path) {
    HKEY hKey;
    if (RegCreateKeyEx(HKEY_CURRENT_USER, kRegKey, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
#if defined(_MSC_VER)
        RegSetValueEx(hKey, kRegValue, 0, REG_SZ,
            reinterpret_cast<const BYTE*>(path.c_str()),
            static_cast<DWORD>((path.size() + 1) * sizeof(wchar_t)));
#else
        RegSetValueEx(hKey, kRegValue, 0, REG_SZ,
            (const BYTE*)path.c_str(), (DWORD)((path.size() + 1) * sizeof(wchar_t)));
#endif
        RegCloseKey(hKey);
    }
}

// Load msstyles path from registry
std::wstring LoadMsStylesPath() {
    HKEY hKey;
    std::wstring path;
    if (RegOpenKeyEx(HKEY_CURRENT_USER, kRegKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        DWORD type = 0, size = 0;
        if (RegQueryValueEx(hKey, kRegValue, NULL, &type, NULL, &size) == ERROR_SUCCESS && type == REG_SZ && size > 2) {
            std::vector<wchar_t> buf(size / sizeof(wchar_t));
            if (RegQueryValueEx(hKey, kRegValue, NULL, &type, reinterpret_cast<LPBYTE>(buf.data()), &size) == ERROR_SUCCESS) {
                path.assign(buf.data());
            }
        }
        RegCloseKey(hKey);
    }
    return path;
}

// Simple file open dialog to choose a .msstyles file
std::wstring PromptForMsStyles(HWND hWnd) {
    wchar_t file[MAX_PATH] = L"";
    OPENFILENAME ofn = { 0 };
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFilter = L"msstyles files\0*.msstyles\0All files\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    ofn.lpstrTitle = L"Select a .msstyles file";
    if (GetOpenFileName(&ofn)) {
        return std::wstring(file);
    }
    return L"";
}

// Add item to TreeView
HTREEITEM AddTreeItem(HWND hTree, HTREEITEM hParent, const std::wstring& text) {
    TVINSERTSTRUCT tvis = {};
    tvis.hParent = hParent;
    tvis.hInsertAfter = TVI_LAST;
    tvis.item.mask = TVIF_TEXT;
    tvis.item.pszText = const_cast<wchar_t*>(text.c_str());
    return TreeView_InsertItem(hTree, &tvis);
}


void ParseCMapBlob(const BYTE* pData, DWORD size, std::map<std::wstring, std::vector<std::wstring>>& groups) 
{ 
    const wchar_t* pStr = reinterpret_cast<const wchar_t*>(pData); DWORD chars = size / sizeof(wchar_t); 
    for (DWORD i = 0; i < chars; ) 
    { 
        if (pStr[i] == L'\0') 
        {
            ++i; continue; 
        }
        std::wstring entry;
        while (i < chars && pStr[i] != L'\0') 
        { 
            entry.push_back(pStr[i]); ++i; 
        }
        if (entry.find(L"DarkMode") != std::wstring::npos)
        { 
            size_t pos = entry.find(L"::"); 
            if (pos != std::wstring::npos) 
            { 
            std::wstring base = entry.substr(pos + 2); groups[base].push_back(entry); 
            } 
            else { groups[L"(unknown)"].push_back(entry);
            }
        } 
    } 
} // Callback for EnumResourceNames 
BOOL CALLBACK EnumNamesProc(HMODULE hModule, LPCWSTR lpszType, LPWSTR lpszName, LONG_PTR lParam)
{ 
    auto* groups = reinterpret_cast<std::map<std::wstring, std::vector<std::wstring>>*>(lParam); 
    // Load resource (default to English language) 
    HRSRC hRes = FindResourceEx(hModule, lpszType, lpszName, MAKELANGID(LANG_ENGLISH, SUBLANG_DEFAULT)); 
    if (!hRes) return TRUE; 
    HGLOBAL hData = LoadResource(hModule, hRes); 
    if (!hData) 
        return TRUE; 
    DWORD size = SizeofResource(hModule, hRes); 
    const BYTE* pData = (const BYTE*)LockResource(hData); 
    ParseCMapBlob(pData, size, *groups); 
    return TRUE; // continue enumeration 
} 
std::map<std::wstring, std::vector<std::wstring>> ParseDarkModeGroups(const std::wstring& path) 
{
    std::map<std::wstring, std::vector<std::wstring>> groups; 
    HMODULE hMod = LoadLibraryEx(path.c_str(), NULL, LOAD_LIBRARY_AS_DATAFILE); 
    if (!hMod) return groups; 
    // Enumerate all CMAP resources instead of assuming ID=1 
    EnumResourceNames(hMod, L"CMap", EnumNamesProc, reinterpret_cast<LONG_PTR>(&groups)); 
   FreeLibrary(hMod); 
return groups;
}
// Populate tree with path root, base classes as children, variants as grandchildren
void PopulateTree(HWND hTree, const std::wstring& path,
    const std::map<std::wstring, std::vector<std::wstring>>& groups) {
    TreeView_DeleteAllItems(hTree);
    HTREEITEM root = AddTreeItem(hTree, TVI_ROOT, path);
    for (const auto& kv : groups) {
        HTREEITEM baseNode = AddTreeItem(hTree, root, kv.first);
        for (const auto& variant : kv.second) {
            AddTreeItem(hTree, baseNode, variant);
        }
    }
    TreeView_Expand(hTree, root, TVE_EXPAND);
}

// Create TreeView control
HWND CreateTreeView(HWND hParent) {
    INITCOMMONCONTROLSEX icex = { sizeof(icex), ICC_TREEVIEW_CLASSES };
    InitCommonControlsEx(&icex);

    RECT rc; GetClientRect(hParent, &rc);
    HWND hTree = CreateWindowEx(WS_EX_CLIENTEDGE, WC_TREEVIEW, L"",
        WS_CHILD | WS_VISIBLE | TVS_HASLINES | TVS_HASBUTTONS | TVS_LINESATROOT,
        0, 0, rc.right - rc.left, rc.bottom - rc.top,
        hParent, NULL, GetModuleHandle(NULL), NULL);
    return hTree;
}

// Resize child controls on window resize
void ResizeChildren(HWND hWnd) {
    if (!g_hTree) return;
    RECT rc; GetClientRect(hWnd, &rc);
    SetWindowPos(g_hTree, NULL, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);
}

// Refresh tree from current path
void RefreshTree(HWND hWnd) {
    if (g_msstylesPath.empty()) return;
    auto groups = ParseDarkModeGroups(g_msstylesPath);
    PopulateTree(g_hTree, g_msstylesPath, groups);
}

// Command handlers
void OnChooseFile(HWND hWnd) {
    std::wstring path = PromptForMsStyles(hWnd);
    if (!path.empty()) {
        g_msstylesPath = path;
        SaveMsStylesPath(g_msstylesPath);
        RefreshTree(hWnd);
    }
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_hTree = CreateTreeView(hWnd);
        // Menu for choosing file
        HMENU hMenu = CreateMenu();
        HMENU hFile = CreatePopupMenu();
        AppendMenu(hFile, MF_STRING, 1001, L"Open .msstyles...");
        AppendMenu(hFile, MF_SEPARATOR, 0, NULL);
        AppendMenu(hFile, MF_STRING, 1002, L"Exit");
        AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFile, L"File");
        SetMenu(hWnd, hMenu);

        // Load path from registry or prompt
        g_msstylesPath = LoadMsStylesPath();
        if (g_msstylesPath.empty()) {
            g_msstylesPath = PromptForMsStyles(hWnd);
            if (!g_msstylesPath.empty()) SaveMsStylesPath(g_msstylesPath);
        }
        if (!g_msstylesPath.empty()) RefreshTree(hWnd);
        break;
    }
    case WM_SIZE:
        ResizeChildren(hWnd);
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case 1001: OnChooseFile(hWnd); break;
        case 1002: PostQuitMessage(0); break;
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int nCmdShow) {
    const wchar_t* kClassName = L"MsstylesCMapViewer";
    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = kClassName;

    if (!RegisterClass(&wc)) return 0;

    HWND hWnd = CreateWindowEx(0, kClassName, L"CMap DarkMode Classes Viewer",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 900, 600,
        NULL, NULL, hInst, NULL);
    if (!hWnd) return 0;

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}
