#pragma once

#include <string>

namespace mytools {

class SecretStoreDpapi {
public:
    bool SmokeTest(std::wstring* error_message) const;
    bool ProtectUtf8ToBase64(const std::string& plain_utf8,
                             std::string* protected_base64,
                             std::wstring* error_message) const;
    bool UnprotectBase64ToUtf8(const std::string& protected_base64,
                               std::string* plain_utf8,
                               std::wstring* error_message) const;
};

}  // namespace mytools
