#include <windows.h>
#include <cstdint>
#include <string>

namespace {

std::wstring Utf8ToWide(const char* utf8) {
    if (utf8 == nullptr) return std::wstring();
    const int needed = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, nullptr, 0);
    if (needed <= 0) return std::wstring();
    std::wstring wide(static_cast<size_t>(needed), L'\0');
    const int written = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wide.data(), needed);
    if (written <= 0) return std::wstring();
    if (!wide.empty() && wide.back() == L'\0') wide.pop_back();
    return wide;
}

} // namespace

extern "C" __declspec(dllexport) uint64_t mmshell_get_file_size_utf8(const char* path_utf8) {
    const std::wstring path = Utf8ToWide(path_utf8);
    if (path.empty()) return 0;
    HANDLE h = CreateFileW(
        path.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );
    if (h == INVALID_HANDLE_VALUE) return 0;
    LARGE_INTEGER li{};
    const BOOL ok = GetFileSizeEx(h, &li);
    CloseHandle(h);
    if (!ok) return 0;
    return static_cast<uint64_t>(li.QuadPart);
}

extern "C" __declspec(dllexport) int mmshell_read_file_chunk_utf8(
    const char* path_utf8,
    uint64_t offset,
    uint32_t max_len,
    uint8_t* out_buf,
    uint32_t* out_read
) {
    if (out_read != nullptr) *out_read = 0;
    if (path_utf8 == nullptr || out_buf == nullptr || out_read == nullptr || max_len == 0) return -1;
    const std::wstring path = Utf8ToWide(path_utf8);
    if (path.empty()) return -2;

    HANDLE h = CreateFileW(
        path.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );
    if (h == INVALID_HANDLE_VALUE) return -3;

    LARGE_INTEGER li{};
    li.QuadPart = static_cast<LONGLONG>(offset);
    if (!SetFilePointerEx(h, li, nullptr, FILE_BEGIN)) {
        CloseHandle(h);
        return -4;
    }

    DWORD bytes_read = 0;
    const BOOL ok = ReadFile(h, out_buf, static_cast<DWORD>(max_len), &bytes_read, nullptr);
    CloseHandle(h);
    if (!ok) return -5;
    *out_read = static_cast<uint32_t>(bytes_read);
    return 0;
}
