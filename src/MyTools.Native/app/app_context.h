#pragma once

#include <string>

#include <windows.h>

namespace mytools {

class AppContext {
public:
    static AppContext Create(HINSTANCE instance);

    HINSTANCE instance() const noexcept { return instance_; }
    const std::wstring& exe_path() const noexcept { return exe_path_; }
    const std::wstring& exe_dir() const noexcept { return exe_dir_; }
    const std::wstring& product_name() const noexcept { return product_name_; }
    const std::wstring& version() const noexcept { return version_; }

private:
    HINSTANCE instance_ = nullptr;
    std::wstring exe_path_;
    std::wstring exe_dir_;
    std::wstring product_name_ = L"AJun Tools Native";
    std::wstring version_ = L"0.1.0-native-shell";
};

}  // namespace mytools
