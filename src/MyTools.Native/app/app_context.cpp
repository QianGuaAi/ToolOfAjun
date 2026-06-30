#include "app/app_context.h"

#include <array>

namespace mytools {
namespace {

std::wstring GetModulePath(HINSTANCE instance) {
    std::array<wchar_t, MAX_PATH> buffer{};
    DWORD size = GetModuleFileNameW(instance, buffer.data(), static_cast<DWORD>(buffer.size()));
    while (size == buffer.size()) {
        std::wstring expanded(buffer.begin(), buffer.end());
        expanded.resize(expanded.size() * 2);
        size = GetModuleFileNameW(instance, expanded.data(), static_cast<DWORD>(expanded.size()));
        if (size < expanded.size()) {
            expanded.resize(size);
            return expanded;
        }
    }

    return std::wstring(buffer.data(), size);
}

std::wstring GetDirectoryName(const std::wstring& path) {
    const size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos) {
        return L".";
    }
    return path.substr(0, slash);
}

}  // namespace

AppContext AppContext::Create(HINSTANCE instance) {
    AppContext context;
    context.instance_ = instance;
    context.exe_path_ = GetModulePath(instance);
    context.exe_dir_ = GetDirectoryName(context.exe_path_);
    return context;
}

}  // namespace mytools
