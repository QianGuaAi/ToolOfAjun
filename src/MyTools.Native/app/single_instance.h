#pragma once

#include <windows.h>

namespace mytools {

class SingleInstance {
public:
    SingleInstance();
    ~SingleInstance();

    SingleInstance(const SingleInstance&) = delete;
    SingleInstance& operator=(const SingleInstance&) = delete;

    bool IsPrimary() const noexcept { return mutex_ != nullptr && !already_exists_; }

private:
    HANDLE mutex_ = nullptr;
    bool already_exists_ = false;
};

}  // namespace mytools
