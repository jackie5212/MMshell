#define UNICODE
#define _UNICODE
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <objidl.h>
#include <shlobj.h>
#include <strsafe.h>
#include <string>
#include <vector>
#include <atomic>
#include <iostream>

namespace {

UINT g_cfFileGroupDescriptorW = 0;
UINT g_cfFileContents = 0;
UINT g_cfPreferredDropEffect = 0;
UINT g_cfPerformedDropEffect = 0;
UINT g_cfPasteSucceeded = 0;

HGLOBAL alloc_hglobal_copy(const void* src, SIZE_T size) {
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, size);
    if (!h) return nullptr;
    void* p = GlobalLock(h);
    if (!p) {
        GlobalFree(h);
        return nullptr;
    }
    memcpy(p, src, size);
    GlobalUnlock(h);
    return h;
}

HGLOBAL build_drop_effect_hglobal(DWORD effect) {
    return alloc_hglobal_copy(&effect, sizeof(effect));
}

HGLOBAL build_file_group_descriptor(const std::wstring& fileName, ULONGLONG fileSize) {
    FILEGROUPDESCRIPTORW fgd{};
    fgd.cItems = 1;
    FILEDESCRIPTORW& fd = fgd.fgd[0];
    fd.dwFlags = FD_FILESIZE | FD_ATTRIBUTES | FD_PROGRESSUI;
    fd.dwFileAttributes = FILE_ATTRIBUTE_NORMAL;
    fd.nFileSizeHigh = static_cast<DWORD>(fileSize >> 32);
    fd.nFileSizeLow = static_cast<DWORD>(fileSize & 0xFFFFFFFFULL);
    StringCchCopyW(fd.cFileName, ARRAYSIZE(fd.cFileName), fileName.c_str());
    return alloc_hglobal_copy(&fgd, sizeof(FILEGROUPDESCRIPTORW));
}

class ProbeDropSource final : public IDropSource {
public:
    ProbeDropSource() : ref_(1) {}
    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDropSource) {
            *ppv = static_cast<IDropSource*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++ref_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG n = --ref_;
        if (n == 0) delete this;
        return n;
    }
    IFACEMETHODIMP QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState) override {
        if (fEscapePressed) return DRAGDROP_S_CANCEL;
        if ((grfKeyState & MK_LBUTTON) == 0) return DRAGDROP_S_DROP;
        return S_OK;
    }
    IFACEMETHODIMP GiveFeedback(DWORD) override { return DRAGDROP_S_USEDEFAULTCURSORS; }
private:
    std::atomic<ULONG> ref_;
};

class ProbeDataObject final : public IDataObject {
public:
    ProbeDataObject(std::wstring name, std::vector<BYTE> bytes)
        : ref_(1), fileName_(std::move(name)), data_(std::move(bytes)) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IDataObject) {
            *ppv = static_cast<IDataObject*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++ref_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG n = --ref_;
        if (n == 0) delete this;
        return n;
    }

    IFACEMETHODIMP GetData(FORMATETC* pFmt, STGMEDIUM* pMed) override {
        if (!pFmt || !pMed) return E_POINTER;
        ZeroMemory(pMed, sizeof(*pMed));
        pMed->tymed = TYMED_HGLOBAL;

        if (pFmt->cfFormat == g_cfFileGroupDescriptorW && (pFmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL h = build_file_group_descriptor(fileName_, static_cast<ULONGLONG>(data_.size()));
            if (!h) return E_OUTOFMEMORY;
            pMed->hGlobal = h;
            return S_OK;
        }

        if (pFmt->cfFormat == g_cfFileContents) {
            if (pFmt->tymed & TYMED_ISTREAM) {
                HGLOBAL h = alloc_hglobal_copy(data_.data(), data_.size());
                if (!h) return E_OUTOFMEMORY;
                IStream* stm = nullptr;
                HRESULT hr = CreateStreamOnHGlobal(h, TRUE, &stm);
                if (FAILED(hr)) {
                    GlobalFree(h);
                    return hr;
                }
                pMed->tymed = TYMED_ISTREAM;
                pMed->pstm = stm;
                return S_OK;
            }
            if (pFmt->tymed & TYMED_HGLOBAL) {
                HGLOBAL h = alloc_hglobal_copy(data_.data(), data_.size());
                if (!h) return E_OUTOFMEMORY;
                pMed->tymed = TYMED_HGLOBAL;
                pMed->hGlobal = h;
                return S_OK;
            }
            return DV_E_TYMED;
        }

        if (pFmt->cfFormat == g_cfPreferredDropEffect && (pFmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL h = build_drop_effect_hglobal(DROPEFFECT_COPY);
            if (!h) return E_OUTOFMEMORY;
            pMed->hGlobal = h;
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    IFACEMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) override { return DATA_E_FORMATETC; }
    IFACEMETHODIMP QueryGetData(FORMATETC* pFmt) override {
        if (!pFmt) return E_POINTER;
        if (pFmt->cfFormat == g_cfFileGroupDescriptorW && (pFmt->tymed & TYMED_HGLOBAL)) return S_OK;
        if (pFmt->cfFormat == g_cfFileContents && (pFmt->tymed & (TYMED_HGLOBAL | TYMED_ISTREAM))) return S_OK;
        if (pFmt->cfFormat == g_cfPreferredDropEffect && (pFmt->tymed & TYMED_HGLOBAL)) return S_OK;
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) override { return E_NOTIMPL; }
    IFACEMETHODIMP SetData(FORMATETC* pFmt, STGMEDIUM*, BOOL) override {
        if (!pFmt) return E_POINTER;
        if (pFmt->cfFormat == g_cfPerformedDropEffect || pFmt->cfFormat == g_cfPasteSucceeded || pFmt->cfFormat == g_cfPreferredDropEffect) {
            return S_OK;
        }
        return S_OK;
    }
    IFACEMETHODIMP EnumFormatEtc(DWORD dir, IEnumFORMATETC** ppEnum) override {
        if (!ppEnum) return E_POINTER;
        *ppEnum = nullptr;
        if (dir != DATADIR_GET) return E_NOTIMPL;
        FORMATETC fmt[4]{};
        fmt[0] = { static_cast<CLIPFORMAT>(g_cfFileGroupDescriptorW), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        fmt[1] = { static_cast<CLIPFORMAT>(g_cfFileContents), nullptr, DVASPECT_CONTENT, 0, TYMED_ISTREAM };
        fmt[2] = { static_cast<CLIPFORMAT>(g_cfFileContents), nullptr, DVASPECT_CONTENT, 0, TYMED_HGLOBAL };
        fmt[3] = { static_cast<CLIPFORMAT>(g_cfPreferredDropEffect), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return SHCreateStdEnumFmtEtc(4, fmt, ppEnum);
    }
    IFACEMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }

private:
    std::atomic<ULONG> ref_;
    std::wstring fileName_;
    std::vector<BYTE> data_;
};

} // namespace

struct ProbeContext {
    IDataObject* dataObj = nullptr;
    IDropSource* dropSource = nullptr;
};

LRESULT CALLBACK ProbeWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    auto* ctx = reinterpret_cast<ProbeContext*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    switch (msg) {
    case WM_CREATE: {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(cs->lpCreateParams));
        return 0;
    }
    case WM_PAINT: {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc{};
        GetClientRect(hwnd, &rc);
        const wchar_t* text = L"按住鼠标左键并拖出窗口到资源管理器文件夹。\n"
                              L"Hold left mouse button and drag out to Explorer.";
        DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_WORDBREAK);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_LBUTTONDOWN: {
        if (!ctx || !ctx->dataObj || !ctx->dropSource) return 0;
        DWORD effect = DROPEFFECT_NONE;
        // 在左键按下事件里触发，确保系统键盘状态与拖拽状态一致。
        HRESULT hr = DoDragDrop(ctx->dataObj, ctx->dropSource, DROPEFFECT_COPY, &effect);
        std::wcout << L"[probe] DoDragDrop hr=0x" << std::hex << hr << L", effect=0x" << effect << std::endl;
        return 0;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

int wmain() {
    HRESULT hr = OleInitialize(nullptr);
    if (FAILED(hr)) {
        std::wcerr << L"OleInitialize failed: 0x" << std::hex << hr << std::endl;
        return 1;
    }

    g_cfFileGroupDescriptorW = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
    g_cfFileContents = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    g_cfPreferredDropEffect = RegisterClipboardFormatW(L"Preferred DropEffect");
    g_cfPerformedDropEffect = RegisterClipboardFormatW(L"Performed DropEffect");
    g_cfPasteSucceeded = RegisterClipboardFormatW(L"Paste Succeeded");

    const std::wstring fileName = L"probe_from_cpp_40mb.bin";
    constexpr size_t kProbeSize = 40u * 1024u * 1024u; // 40 MB
    std::vector<BYTE> bytes(kProbeSize);
    for (size_t i = 0; i < bytes.size(); ++i) {
        bytes[i] = static_cast<BYTE>(i & 0xFF);
    }

    ProbeContext ctx;
    ctx.dataObj = new ProbeDataObject(fileName, bytes);
    ctx.dropSource = new ProbeDropSource();

    HINSTANCE hinst = GetModuleHandleW(nullptr);
    const wchar_t* clsName = L"NativeDragProbeWnd";
    WNDCLASSW wc{};
    wc.lpfnWndProc = ProbeWndProc;
    wc.hInstance = hinst;
    wc.lpszClassName = clsName;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    RegisterClassW(&wc);

    HWND hwnd = CreateWindowExW(
        0, clsName, L"Native Drag Probe (C++)",
        WS_OVERLAPPEDWINDOW | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 560, 220,
        nullptr, nullptr, hinst, &ctx
    );
    if (!hwnd) {
        std::wcerr << L"CreateWindow failed." << std::endl;
        ctx.dropSource->Release();
        ctx.dataObj->Release();
        OleUninitialize();
        return 2;
    }

    std::wcout << L"[probe] window ready. Hold left mouse and drag out." << std::endl;
    MSG m{};
    while (GetMessageW(&m, nullptr, 0, 0) > 0) {
        TranslateMessage(&m);
        DispatchMessageW(&m);
    }

    ctx.dropSource->Release();
    ctx.dataObj->Release();
    OleUninitialize();
    return 0;
}
