#define UNICODE
#define _UNICODE
#define NOMINMAX
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <objidl.h>
#include <shlobj.h>
#include <exdisp.h>
#include <shlwapi.h>
#include <strsafe.h>
#include <string>
#include <vector>
#include <atomic>
#include <thread>
#include <chrono>

namespace {

UINT g_cfFileGroupDescriptorW = 0;
UINT g_cfFileContents = 0;
UINT g_cfPreferredDropEffect = 0;
UINT g_cfPerformedDropEffect = 0;
UINT g_cfPasteSucceeded = 0;

std::wstring Utf8ToWide(const char* utf8) {
    if (!utf8) return std::wstring();
    const int need = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    if (need <= 0) return std::wstring();
    std::wstring out(static_cast<size_t>(need), L'\0');
    const int wrote = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, out.data(), need);
    if (wrote <= 0) return std::wstring();
    if (!out.empty() && out.back() == L'\0') out.pop_back();
    return out;
}

std::string WideToUtf8(const std::wstring& wide) {
    if (wide.empty()) return std::string();
    const int need = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (need <= 0) return std::string();
    std::string out(static_cast<size_t>(need), '\0');
    const int wrote = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, out.data(), need, nullptr, nullptr);
    if (wrote <= 0) return std::string();
    if (!out.empty() && out.back() == '\0') out.pop_back();
    return out;
}

std::wstring UrlDecode(const std::wstring& input) {
    std::wstring out;
    out.reserve(input.size());
    for (size_t i = 0; i < input.size(); ++i) {
        const wchar_t c = input[i];
        if (c == L'%' && i + 2 < input.size()) {
            auto hex = [](wchar_t ch) -> int {
                if (ch >= L'0' && ch <= L'9') return ch - L'0';
                if (ch >= L'a' && ch <= L'f') return ch - L'a' + 10;
                if (ch >= L'A' && ch <= L'F') return ch - L'A' + 10;
                return -1;
            };
            int hi = hex(input[i + 1]);
            int lo = hex(input[i + 2]);
            if (hi >= 0 && lo >= 0) {
                out.push_back(static_cast<wchar_t>((hi << 4) | lo));
                i += 2;
                continue;
            }
        }
        out.push_back(c);
    }
    return out;
}

std::wstring FileUrlToPath(const std::wstring& url) {
    std::wstring u = url;
    if (u.rfind(L"file:///", 0) == 0) {
        u = u.substr(8);
    } else if (u.rfind(L"file://", 0) == 0) {
        u = u.substr(7);
    }
    u = UrlDecode(u);
    for (auto& ch : u) {
        if (ch == L'/') ch = L'\\';
    }
    return u;
}

bool GetDesktopFolderPath(std::wstring& outPath) {
    PWSTR path = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_Desktop, 0, nullptr, &path);
    if (FAILED(hr) || !path) return false;
    outPath.assign(path);
    CoTaskMemFree(path);
    return !outPath.empty();
}

bool TryGetExplorerFolderFromHwnd(HWND targetRoot, std::wstring& outPath) {
    IShellWindows* shellWindows = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_PPV_ARGS(&shellWindows));
    if (FAILED(hr) || !shellWindows) return false;
    bool found = false;
    long count = 0;
    if (SUCCEEDED(shellWindows->get_Count(&count))) {
        for (long i = 0; i < count; ++i) {
            VARIANT v{};
            V_VT(&v) = VT_I4;
            V_I4(&v) = i;
            IDispatch* disp = nullptr;
            if (FAILED(shellWindows->Item(v, &disp)) || !disp) continue;
            IWebBrowserApp* web = nullptr;
            if (SUCCEEDED(disp->QueryInterface(IID_PPV_ARGS(&web))) && web) {
                SHANDLE_PTR hwndVal = 0;
                if (SUCCEEDED(web->get_HWND(&hwndVal))) {
                    HWND w = reinterpret_cast<HWND>(hwndVal);
                    HWND root = GetAncestor(w, GA_ROOT);
                    if (root == targetRoot) {
                        BSTR locationUrl = nullptr;
                        if (SUCCEEDED(web->get_LocationURL(&locationUrl)) && locationUrl) {
                            std::wstring url(locationUrl, SysStringLen(locationUrl));
                            SysFreeString(locationUrl);
                            std::wstring path = FileUrlToPath(url);
                            if (!path.empty()) {
                                outPath = path;
                                found = true;
                            }
                        }
                    }
                }
                web->Release();
            }
            disp->Release();
            if (found) break;
        }
    }
    shellWindows->Release();
    return found;
}

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

HGLOBAL build_hdrop_for_path(const std::wstring& inPath) {
    wchar_t full[MAX_PATH * 4]{};
    DWORD n = GetFullPathNameW(inPath.c_str(), static_cast<DWORD>(std::size(full)), full, nullptr);
    if (n == 0 || n >= std::size(full)) return nullptr;
    std::wstring path(full);
    // DROPFILES + UTF-16 path + double-null terminator
    const SIZE_T header = sizeof(DROPFILES);
    const SIZE_T chars = path.size() + 2;
    const SIZE_T bytes = header + chars * sizeof(wchar_t);
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes);
    if (!h) return nullptr;
    BYTE* p = static_cast<BYTE*>(GlobalLock(h));
    if (!p) {
        GlobalFree(h);
        return nullptr;
    }
    auto* df = reinterpret_cast<DROPFILES*>(p);
    df->pFiles = static_cast<DWORD>(header);
    df->fWide = TRUE;
    wchar_t* dst = reinterpret_cast<wchar_t*>(p + header);
    memcpy(dst, path.c_str(), path.size() * sizeof(wchar_t));
    dst[path.size()] = L'\0';
    dst[path.size() + 1] = L'\0';
    GlobalUnlock(h);
    return h;
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

class DropSource final : public IDropSource {
public:
    DropSource() : ref_(1) {}
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
        if (!n) delete this;
        return n;
    }
    IFACEMETHODIMP QueryContinueDrag(BOOL esc, DWORD key) override {
        if (esc) return DRAGDROP_S_CANCEL;
        if ((key & MK_LBUTTON) == 0) return DRAGDROP_S_DROP;
        return S_OK;
    }
    IFACEMETHODIMP GiveFeedback(DWORD) override { return DRAGDROP_S_USEDEFAULTCURSORS; }
private:
    std::atomic<ULONG> ref_;
};

class HdropDataObject final : public IDataObject {
public:
    explicit HdropDataObject(std::wstring path) : ref_(1), path_(std::move(path)) {}

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
        if (!n) delete this;
        return n;
    }

    IFACEMETHODIMP GetData(FORMATETC* fmt, STGMEDIUM* med) override {
        if (!fmt || !med) return E_POINTER;
        ZeroMemory(med, sizeof(*med));
        med->tymed = TYMED_HGLOBAL;

        if (fmt->cfFormat == CF_HDROP && (fmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL hg = build_hdrop_for_path(path_);
            if (!hg) return E_OUTOFMEMORY;
            med->hGlobal = hg;
            return S_OK;
        }
        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL hg = build_drop_effect_hglobal(DROPEFFECT_COPY);
            if (!hg) return E_OUTOFMEMORY;
            med->hGlobal = hg;
            return S_OK;
        }
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) override { return DATA_E_FORMATETC; }
    IFACEMETHODIMP QueryGetData(FORMATETC* fmt) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == CF_HDROP && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) override { return E_NOTIMPL; }
    IFACEMETHODIMP SetData(FORMATETC* fmt, STGMEDIUM*, BOOL) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == g_cfPerformedDropEffect || fmt->cfFormat == g_cfPasteSucceeded || fmt->cfFormat == g_cfPreferredDropEffect) return S_OK;
        return S_OK;
    }
    IFACEMETHODIMP EnumFormatEtc(DWORD dir, IEnumFORMATETC** ppEnum) override {
        if (!ppEnum) return E_POINTER;
        *ppEnum = nullptr;
        if (dir != DATADIR_GET) return E_NOTIMPL;
        FORMATETC fmts[2]{};
        fmts[0] = { static_cast<CLIPFORMAT>(CF_HDROP), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        fmts[1] = { static_cast<CLIPFORMAT>(g_cfPreferredDropEffect), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return SHCreateStdEnumFmtEtc(2, fmts, ppEnum);
    }
    IFACEMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }

private:
    std::atomic<ULONG> ref_;
    std::wstring path_;
};

class GrowingFileStream final : public IStream {
public:
    GrowingFileStream(std::wstring path, std::wstring donePath, std::wstring errPath, DWORD waitMs)
        : ref_(1), path_(std::move(path)), donePath_(std::move(donePath)), errPath_(std::move(errPath)), waitMs_(waitMs), pos_(0) {}

    IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IStream || riid == IID_ISequentialStream) {
            *ppv = static_cast<IStream*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    IFACEMETHODIMP_(ULONG) AddRef() override { return ++ref_; }
    IFACEMETHODIMP_(ULONG) Release() override {
        ULONG n = --ref_;
        if (!n) delete this;
        return n;
    }

    IFACEMETHODIMP Read(void* pv, ULONG cb, ULONG* pcbRead) override {
        if (!pv) return E_POINTER;
        if (pcbRead) *pcbRead = 0;
        const auto start = GetTickCount64();
        for (;;) {
            if (GetFileAttributesW(errPath_.c_str()) != INVALID_FILE_ATTRIBUTES) {
                // 下载任务已失败，返回空读，避免灾难性 COM 错误。
                return S_OK;
            }
            HANDLE h = CreateFileW(path_.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                   nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (h != INVALID_HANDLE_VALUE) {
                LARGE_INTEGER sz{};
                if (GetFileSizeEx(h, &sz)) {
                    if (pos_ < static_cast<ULONGLONG>(sz.QuadPart)) {
                        LARGE_INTEGER off{};
                        off.QuadPart = static_cast<LONGLONG>(pos_);
                        if (SetFilePointerEx(h, off, nullptr, FILE_BEGIN)) {
                            DWORD rd = 0;
                            DWORD want = static_cast<DWORD>(cb);
                            if (want > static_cast<DWORD>(sz.QuadPart - static_cast<LONGLONG>(pos_))) {
                                want = static_cast<DWORD>(sz.QuadPart - static_cast<LONGLONG>(pos_));
                            }
                            if (ReadFile(h, pv, want, &rd, nullptr)) {
                                CloseHandle(h);
                                pos_ += rd;
                                if (pcbRead) *pcbRead = rd;
                                return S_OK;
                            }
                        }
                    }
                }
                CloseHandle(h);
            }

            const bool done = GetFileAttributesW(donePath_.c_str()) != INVALID_FILE_ATTRIBUTES;
            if (done) {
                if (pcbRead) *pcbRead = 0;
                return S_OK;
            }
            if ((GetTickCount64() - start) > waitMs_) {
                if (pcbRead) *pcbRead = 0;
                return S_OK;
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(50));
        }
    }

    IFACEMETHODIMP Write(const void*, ULONG, ULONG*) override { return E_NOTIMPL; }
    IFACEMETHODIMP Seek(LARGE_INTEGER d, DWORD origin, ULARGE_INTEGER* newPos) override {
        LONGLONG base = 0;
        if (origin == STREAM_SEEK_SET) base = 0;
        else if (origin == STREAM_SEEK_CUR) base = static_cast<LONGLONG>(pos_);
        else if (origin == STREAM_SEEK_END) {
            WIN32_FILE_ATTRIBUTE_DATA fad{};
            if (GetFileAttributesExW(path_.c_str(), GetFileExInfoStandard, &fad)) {
                ULONGLONG s = (static_cast<ULONGLONG>(fad.nFileSizeHigh) << 32) | fad.nFileSizeLow;
                base = static_cast<LONGLONG>(s);
            }
        } else return STG_E_INVALIDFUNCTION;
        LONGLONG p = base + d.QuadPart;
        if (p < 0) p = 0;
        pos_ = static_cast<ULONGLONG>(p);
        if (newPos) newPos->QuadPart = pos_;
        return S_OK;
    }
    IFACEMETHODIMP SetSize(ULARGE_INTEGER) override { return E_NOTIMPL; }
    IFACEMETHODIMP CopyTo(IStream*, ULARGE_INTEGER, ULARGE_INTEGER*, ULARGE_INTEGER*) override { return E_NOTIMPL; }
    IFACEMETHODIMP Commit(DWORD) override { return S_OK; }
    IFACEMETHODIMP Revert() override { return E_NOTIMPL; }
    IFACEMETHODIMP LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    IFACEMETHODIMP UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override { return E_NOTIMPL; }
    IFACEMETHODIMP Stat(STATSTG* p, DWORD) override {
        if (!p) return E_POINTER;
        ZeroMemory(p, sizeof(*p));
        p->type = STGTY_STREAM;
        WIN32_FILE_ATTRIBUTE_DATA fad{};
        if (GetFileAttributesExW(path_.c_str(), GetFileExInfoStandard, &fad)) {
            ULONGLONG s = (static_cast<ULONGLONG>(fad.nFileSizeHigh) << 32) | fad.nFileSizeLow;
            p->cbSize.QuadPart = s;
        }
        return S_OK;
    }
    IFACEMETHODIMP Clone(IStream** ppstm) override {
        if (!ppstm) return E_POINTER;
        auto* s = new GrowingFileStream(path_, donePath_, errPath_, waitMs_);
        s->pos_ = pos_;
        *ppstm = s;
        return S_OK;
    }
private:
    std::atomic<ULONG> ref_;
    std::wstring path_;
    std::wstring donePath_;
    std::wstring errPath_;
    DWORD waitMs_;
    ULONGLONG pos_;
};

class StreamingFileDataObject final : public IDataObject {
public:
    StreamingFileDataObject(std::wstring path, std::wstring name, std::wstring donePath, std::wstring errPath, DWORD waitMs)
        : ref_(1), path_(std::move(path)), name_(std::move(name)), donePath_(std::move(donePath)), errPath_(std::move(errPath)), waitMs_(waitMs) {}

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
        if (!n) delete this;
        return n;
    }

    IFACEMETHODIMP GetData(FORMATETC* fmt, STGMEDIUM* med) override {
        if (!fmt || !med) return E_POINTER;
        ZeroMemory(med, sizeof(*med));
        if (fmt->cfFormat == g_cfFileGroupDescriptorW && (fmt->tymed & TYMED_HGLOBAL)) {
            // 文件大小未知（流式），用 0 + 进度 UI。
            HGLOBAL hg = build_file_group_descriptor(name_, 0);
            if (!hg) return E_OUTOFMEMORY;
            med->tymed = TYMED_HGLOBAL;
            med->hGlobal = hg;
            return S_OK;
        }
        if (fmt->cfFormat == g_cfFileContents && (fmt->tymed & TYMED_ISTREAM)) {
            auto* stm = new GrowingFileStream(path_, donePath_, errPath_, waitMs_);
            med->tymed = TYMED_ISTREAM;
            med->pstm = stm;
            return S_OK;
        }
        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL hg = build_drop_effect_hglobal(DROPEFFECT_COPY);
            if (!hg) return E_OUTOFMEMORY;
            med->tymed = TYMED_HGLOBAL;
            med->hGlobal = hg;
            return S_OK;
        }
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) override { return DATA_E_FORMATETC; }
    IFACEMETHODIMP QueryGetData(FORMATETC* fmt) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == g_cfFileGroupDescriptorW && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        if (fmt->cfFormat == g_cfFileContents && (fmt->tymed & TYMED_ISTREAM)) return S_OK;
        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) override { return E_NOTIMPL; }
    IFACEMETHODIMP SetData(FORMATETC* fmt, STGMEDIUM*, BOOL) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == g_cfPerformedDropEffect || fmt->cfFormat == g_cfPasteSucceeded || fmt->cfFormat == g_cfPreferredDropEffect) return S_OK;
        return S_OK;
    }
    IFACEMETHODIMP EnumFormatEtc(DWORD dir, IEnumFORMATETC** ppEnum) override {
        if (!ppEnum) return E_POINTER;
        *ppEnum = nullptr;
        if (dir != DATADIR_GET) return E_NOTIMPL;
        FORMATETC fmts[3]{};
        fmts[0] = { static_cast<CLIPFORMAT>(g_cfFileGroupDescriptorW), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        fmts[1] = { static_cast<CLIPFORMAT>(g_cfFileContents), nullptr, DVASPECT_CONTENT, 0, TYMED_ISTREAM };
        fmts[2] = { static_cast<CLIPFORMAT>(g_cfPreferredDropEffect), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return SHCreateStdEnumFmtEtc(3, fmts, ppEnum);
    }
    IFACEMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }
private:
    std::atomic<ULONG> ref_;
    std::wstring path_;
    std::wstring name_;
    std::wstring donePath_;
    std::wstring errPath_;
    DWORD waitMs_;
};

class FileDataObject final : public IDataObject {
public:
    FileDataObject(std::wstring path, std::wstring name)
        : ref_(1), path_(std::move(path)), name_(std::move(name)) {}

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
        if (!n) delete this;
        return n;
    }

    IFACEMETHODIMP GetData(FORMATETC* fmt, STGMEDIUM* med) override {
        if (!fmt || !med) return E_POINTER;
        ZeroMemory(med, sizeof(*med));

        if (fmt->cfFormat == g_cfFileGroupDescriptorW && (fmt->tymed & TYMED_HGLOBAL)) {
            HANDLE h = CreateFileW(path_.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                   nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (h == INVALID_HANDLE_VALUE) return E_FAIL;
            LARGE_INTEGER li{};
            BOOL ok = GetFileSizeEx(h, &li);
            CloseHandle(h);
            if (!ok) return E_FAIL;
            HGLOBAL hg = build_file_group_descriptor(name_, static_cast<ULONGLONG>(li.QuadPart));
            if (!hg) return E_OUTOFMEMORY;
            med->tymed = TYMED_HGLOBAL;
            med->hGlobal = hg;
            return S_OK;
        }

        if (fmt->cfFormat == g_cfFileContents) {
            HANDLE h = CreateFileW(path_.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                   nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
            if (h == INVALID_HANDLE_VALUE) return E_FAIL;
            LARGE_INTEGER li{};
            if (!GetFileSizeEx(h, &li)) {
                CloseHandle(h);
                return E_FAIL;
            }
            std::vector<BYTE> buf(static_cast<size_t>(li.QuadPart));
            DWORD total = 0;
            while (total < buf.size()) {
                DWORD chunk = 0;
                if (!ReadFile(h, buf.data() + total, static_cast<DWORD>(buf.size() - total), &chunk, nullptr)) {
                    CloseHandle(h);
                    return E_FAIL;
                }
                if (chunk == 0) break;
                total += chunk;
            }
            CloseHandle(h);
            buf.resize(total);

            if (fmt->tymed & TYMED_ISTREAM) {
                HGLOBAL hg = alloc_hglobal_copy(buf.data(), buf.size());
                if (!hg) return E_OUTOFMEMORY;
                IStream* stm = nullptr;
                HRESULT hr = CreateStreamOnHGlobal(hg, TRUE, &stm);
                if (FAILED(hr)) {
                    GlobalFree(hg);
                    return hr;
                }
                med->tymed = TYMED_ISTREAM;
                med->pstm = stm;
                return S_OK;
            }
            if (fmt->tymed & TYMED_HGLOBAL) {
                HGLOBAL hg = alloc_hglobal_copy(buf.data(), buf.size());
                if (!hg) return E_OUTOFMEMORY;
                med->tymed = TYMED_HGLOBAL;
                med->hGlobal = hg;
                return S_OK;
            }
            return DV_E_TYMED;
        }

        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) {
            HGLOBAL hg = build_drop_effect_hglobal(DROPEFFECT_COPY);
            if (!hg) return E_OUTOFMEMORY;
            med->tymed = TYMED_HGLOBAL;
            med->hGlobal = hg;
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    IFACEMETHODIMP GetDataHere(FORMATETC*, STGMEDIUM*) override { return DATA_E_FORMATETC; }
    IFACEMETHODIMP QueryGetData(FORMATETC* fmt) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == g_cfFileGroupDescriptorW && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        if (fmt->cfFormat == g_cfFileContents && (fmt->tymed & (TYMED_ISTREAM | TYMED_HGLOBAL))) return S_OK;
        if (fmt->cfFormat == g_cfPreferredDropEffect && (fmt->tymed & TYMED_HGLOBAL)) return S_OK;
        return DV_E_FORMATETC;
    }
    IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC*, FORMATETC*) override { return E_NOTIMPL; }
    IFACEMETHODIMP SetData(FORMATETC* fmt, STGMEDIUM*, BOOL) override {
        if (!fmt) return E_POINTER;
        if (fmt->cfFormat == g_cfPerformedDropEffect || fmt->cfFormat == g_cfPasteSucceeded || fmt->cfFormat == g_cfPreferredDropEffect) return S_OK;
        return S_OK;
    }
    IFACEMETHODIMP EnumFormatEtc(DWORD dir, IEnumFORMATETC** ppEnum) override {
        if (!ppEnum) return E_POINTER;
        *ppEnum = nullptr;
        if (dir != DATADIR_GET) return E_NOTIMPL;
        FORMATETC fmts[4]{};
        fmts[0] = { static_cast<CLIPFORMAT>(g_cfFileGroupDescriptorW), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        fmts[1] = { static_cast<CLIPFORMAT>(g_cfFileContents), nullptr, DVASPECT_CONTENT, 0, TYMED_ISTREAM };
        fmts[2] = { static_cast<CLIPFORMAT>(g_cfFileContents), nullptr, DVASPECT_CONTENT, 0, TYMED_HGLOBAL };
        fmts[3] = { static_cast<CLIPFORMAT>(g_cfPreferredDropEffect), nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return SHCreateStdEnumFmtEtc(4, fmts, ppEnum);
    }
    IFACEMETHODIMP DAdvise(FORMATETC*, DWORD, IAdviseSink*, DWORD*) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP DUnadvise(DWORD) override { return OLE_E_ADVISENOTSUPPORTED; }
    IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA**) override { return OLE_E_ADVISENOTSUPPORTED; }

private:
    std::atomic<ULONG> ref_;
    std::wstring path_;
    std::wstring name_;
};

} // namespace

extern "C" __declspec(dllexport) int mmshell_start_virtual_drag_from_file_utf8(
    const char* local_path_utf8,
    const char* display_name_utf8,
    unsigned int* out_effect
) {
    if (!local_path_utf8 || !display_name_utf8) return static_cast<int>(E_INVALIDARG);
    if (out_effect) *out_effect = 0;
    std::wstring path = Utf8ToWide(local_path_utf8);
    std::wstring name = Utf8ToWide(display_name_utf8);
    if (path.empty() || name.empty()) return static_cast<int>(E_INVALIDARG);

    if (!g_cfFileGroupDescriptorW) g_cfFileGroupDescriptorW = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
    if (!g_cfFileContents) g_cfFileContents = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    if (!g_cfPreferredDropEffect) g_cfPreferredDropEffect = RegisterClipboardFormatW(L"Preferred DropEffect");
    if (!g_cfPerformedDropEffect) g_cfPerformedDropEffect = RegisterClipboardFormatW(L"Performed DropEffect");
    if (!g_cfPasteSucceeded) g_cfPasteSucceeded = RegisterClipboardFormatW(L"Paste Succeeded");

    HRESULT hr = OleInitialize(nullptr);
    if (FAILED(hr)) return static_cast<int>(hr);

    IDataObject* obj = new FileDataObject(path, name);
    IDropSource* src = new DropSource();
    DWORD effect = DROPEFFECT_NONE;
    hr = DoDragDrop(obj, src, DROPEFFECT_COPY, &effect);
    if (out_effect) *out_effect = effect;
    src->Release();
    obj->Release();
    OleUninitialize();
    return static_cast<int>(hr);
}

extern "C" __declspec(dllexport) int mmshell_start_hdrop_drag_from_file_utf8(
    const char* local_path_utf8,
    unsigned int* out_effect
) {
    if (!local_path_utf8) return static_cast<int>(E_INVALIDARG);
    if (out_effect) *out_effect = 0;
    std::wstring path = Utf8ToWide(local_path_utf8);
    if (path.empty()) return static_cast<int>(E_INVALIDARG);

    if (!g_cfPreferredDropEffect) g_cfPreferredDropEffect = RegisterClipboardFormatW(L"Preferred DropEffect");
    if (!g_cfPerformedDropEffect) g_cfPerformedDropEffect = RegisterClipboardFormatW(L"Performed DropEffect");
    if (!g_cfPasteSucceeded) g_cfPasteSucceeded = RegisterClipboardFormatW(L"Paste Succeeded");

    HRESULT hr = OleInitialize(nullptr);
    if (FAILED(hr)) return static_cast<int>(hr);

    IDataObject* obj = new HdropDataObject(path);
    IDropSource* src = new DropSource();
    DWORD effect = DROPEFFECT_NONE;
    hr = DoDragDrop(obj, src, DROPEFFECT_COPY, &effect);
    if (out_effect) *out_effect = effect;
    src->Release();
    obj->Release();
    OleUninitialize();
    return static_cast<int>(hr);
}

extern "C" __declspec(dllexport) int mmshell_start_virtual_drag_streaming_utf8(
    const char* local_path_utf8,
    const char* display_name_utf8,
    const char* done_marker_utf8,
    const char* err_marker_utf8,
    unsigned int wait_ms,
    unsigned int* out_effect
) {
    if (!local_path_utf8 || !display_name_utf8 || !done_marker_utf8 || !err_marker_utf8) return static_cast<int>(E_INVALIDARG);
    if (out_effect) *out_effect = 0;
    std::wstring path = Utf8ToWide(local_path_utf8);
    std::wstring name = Utf8ToWide(display_name_utf8);
    std::wstring donePath = Utf8ToWide(done_marker_utf8);
    std::wstring errPath = Utf8ToWide(err_marker_utf8);
    if (path.empty() || name.empty() || donePath.empty() || errPath.empty()) return static_cast<int>(E_INVALIDARG);

    if (!g_cfFileGroupDescriptorW) g_cfFileGroupDescriptorW = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
    if (!g_cfFileContents) g_cfFileContents = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    if (!g_cfPreferredDropEffect) g_cfPreferredDropEffect = RegisterClipboardFormatW(L"Preferred DropEffect");
    if (!g_cfPerformedDropEffect) g_cfPerformedDropEffect = RegisterClipboardFormatW(L"Performed DropEffect");
    if (!g_cfPasteSucceeded) g_cfPasteSucceeded = RegisterClipboardFormatW(L"Paste Succeeded");

    HRESULT hr = OleInitialize(nullptr);
    if (FAILED(hr)) return static_cast<int>(hr);

    IDataObject* obj = new StreamingFileDataObject(path, name, donePath, errPath, wait_ms ? wait_ms : 300000);
    IDropSource* src = new DropSource();
    DWORD effect = DROPEFFECT_NONE;
    hr = DoDragDrop(obj, src, DROPEFFECT_COPY, &effect);
    if (out_effect) *out_effect = effect;
    src->Release();
    obj->Release();
    OleUninitialize();
    return static_cast<int>(hr);
}

extern "C" __declspec(dllexport) int mmshell_wait_mouse_release_target_utf8(
    char* out_path_utf8,
    unsigned int out_path_capacity,
    unsigned int* out_effect
) {
    if (!out_path_utf8 || out_path_capacity == 0) return static_cast<int>(E_INVALIDARG);
    out_path_utf8[0] = '\0';
    if (out_effect) *out_effect = 0;

    // 等待左键释放，模拟“拖拽意图结束”时刻。
    while ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0) {
        Sleep(10);
    }

    POINT pt{};
    if (!GetCursorPos(&pt)) return static_cast<int>(E_FAIL);
    HWND hwnd = WindowFromPoint(pt);
    if (!hwnd) return static_cast<int>(E_FAIL);
    HWND root = GetAncestor(hwnd, GA_ROOT);
    if (!root) root = hwnd;

    wchar_t cls[128]{};
    GetClassNameW(root, cls, static_cast<int>(std::size(cls)));
    std::wstring className(cls);

    HRESULT hrCom = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    bool inited = SUCCEEDED(hrCom);
    std::wstring path;
    bool ok = false;
    if (className == L"Progman" || className == L"WorkerW") {
        ok = GetDesktopFolderPath(path);
    } else {
        ok = TryGetExplorerFolderFromHwnd(root, path);
        if (!ok) {
            HWND fg = GetForegroundWindow();
            if (fg) {
                ok = TryGetExplorerFolderFromHwnd(GetAncestor(fg, GA_ROOT), path);
            }
        }
    }
    if (inited) CoUninitialize();

    if (!ok || path.empty()) return static_cast<int>(S_FALSE);
    std::string utf8 = WideToUtf8(path);
    if (utf8.empty()) return static_cast<int>(S_FALSE);
    if (utf8.size() + 1 > out_path_capacity) return static_cast<int>(E_FAIL);
    memcpy(out_path_utf8, utf8.c_str(), utf8.size() + 1);
    if (out_effect) *out_effect = 1; // copy
    return static_cast<int>(S_OK);
}

extern "C" __declspec(dllexport) int mmshell_detect_drop_target_utf8(
    char* out_path_utf8,
    unsigned int out_path_capacity,
    unsigned int* out_effect
) {
    if (!out_path_utf8 || out_path_capacity == 0) return static_cast<int>(E_INVALIDARG);
    out_path_utf8[0] = '\0';
    if (out_effect) *out_effect = 0;

    POINT pt{};
    if (!GetCursorPos(&pt)) return static_cast<int>(E_FAIL);
    HWND hwnd = WindowFromPoint(pt);
    if (!hwnd) return static_cast<int>(E_FAIL);
    HWND root = GetAncestor(hwnd, GA_ROOT);
    if (!root) root = hwnd;

    wchar_t cls[128]{};
    GetClassNameW(root, cls, static_cast<int>(std::size(cls)));
    std::wstring className(cls);

    HRESULT hrCom = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    bool inited = SUCCEEDED(hrCom);
    std::wstring path;
    bool ok = false;
    if (className == L"Progman" || className == L"WorkerW") {
        ok = GetDesktopFolderPath(path);
    } else {
        ok = TryGetExplorerFolderFromHwnd(root, path);
        if (!ok) {
            HWND fg = GetForegroundWindow();
            if (fg) {
                ok = TryGetExplorerFolderFromHwnd(GetAncestor(fg, GA_ROOT), path);
            }
        }
    }
    if (inited) CoUninitialize();

    if (!ok || path.empty()) return static_cast<int>(S_FALSE);
    std::string utf8 = WideToUtf8(path);
    if (utf8.empty()) return static_cast<int>(S_FALSE);
    if (utf8.size() + 1 > out_path_capacity) return static_cast<int>(E_FAIL);
    memcpy(out_path_utf8, utf8.c_str(), utf8.size() + 1);
    if (out_effect) *out_effect = 1;
    return static_cast<int>(S_OK);
}
