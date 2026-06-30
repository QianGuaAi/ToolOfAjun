#include "app/single_instance.h"

namespace mytools {

SingleInstance::SingleInstance() {
    mutex_ = CreateMutexW(nullptr, TRUE, L"Local\\AJunTools.MyToolsNative.SingleInstance");
    already_exists_ = GetLastError() == ERROR_ALREADY_EXISTS;
}

SingleInstance::~SingleInstance() {
    if (mutex_ != nullptr) {
        if (!already_exists_) {
            ReleaseMutex(mutex_);
        }
        CloseHandle(mutex_);
        mutex_ = nullptr;
    }
}

}  // namespace mytools
