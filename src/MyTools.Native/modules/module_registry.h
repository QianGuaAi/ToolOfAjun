#pragma once

#include <string>
#include <vector>

namespace mytools {

enum class ModuleId {
    Home,
    CodexProfiles,
};

struct ModuleInfo {
    ModuleId id = ModuleId::Home;
    std::wstring title;
    std::wstring subtitle;
    std::wstring status;
    std::vector<std::wstring> bullets;
};

ModuleInfo BuildHomeModuleInfo();

}  // namespace mytools
