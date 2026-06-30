#include "services/secret_store_dpapi.h"

#include <string>
#include <vector>

#include <windows.h>
#include <dpapi.h>
#include <wincrypt.h>

namespace mytools {
namespace {

DATA_BLOB BlobFromString(const std::wstring& text) {
    DATA_BLOB blob{};
    blob.pbData = reinterpret_cast<BYTE*>(const_cast<wchar_t*>(text.data()));
    blob.cbData = static_cast<DWORD>((text.size() + 1) * sizeof(wchar_t));
    return blob;
}

DATA_BLOB BlobFromBytes(const std::string& text) {
    DATA_BLOB blob{};
    blob.pbData = reinterpret_cast<BYTE*>(const_cast<char*>(text.data()));
    blob.cbData = static_cast<DWORD>(text.size());
    return blob;
}

std::wstring StringFromBlob(const DATA_BLOB& blob) {
    if (blob.pbData == nullptr || blob.cbData < sizeof(wchar_t)) {
        return {};
    }
    return reinterpret_cast<const wchar_t*>(blob.pbData);
}

void FreeBlob(DATA_BLOB* blob) {
    if (blob != nullptr && blob->pbData != nullptr) {
        SecureZeroMemory(blob->pbData, blob->cbData);
        LocalFree(blob->pbData);
        blob->pbData = nullptr;
        blob->cbData = 0;
    }
}

bool DecodeBase64(const std::string& text, std::vector<BYTE>* bytes) {
    if (bytes == nullptr || text.empty()) {
        return false;
    }

    DWORD required = 0;
    if (!CryptStringToBinaryA(text.c_str(),
                              static_cast<DWORD>(text.size()),
                              CRYPT_STRING_BASE64 | CRYPT_STRING_STRICT,
                              nullptr,
                              &required,
                              nullptr,
                              nullptr) ||
        required == 0) {
        return false;
    }

    bytes->assign(required, 0);
    return CryptStringToBinaryA(text.c_str(),
                                static_cast<DWORD>(text.size()),
                                CRYPT_STRING_BASE64 | CRYPT_STRING_STRICT,
                                bytes->data(),
                                &required,
                                nullptr,
                                nullptr) != FALSE;
}

bool EncodeBase64(const BYTE* bytes, DWORD size, std::string* text) {
    if (text == nullptr || bytes == nullptr || size == 0) {
        return false;
    }

    DWORD required = 0;
    if (!CryptBinaryToStringA(bytes,
                              size,
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              nullptr,
                              &required) ||
        required == 0) {
        return false;
    }

    text->assign(required, '\0');
    if (!CryptBinaryToStringA(bytes,
                              size,
                              CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF,
                              text->data(),
                              &required)) {
        text->clear();
        return false;
    }

    if (required > 0 && !text->empty() && text->back() == '\0') {
        text->resize(required - 1);
    } else {
        text->resize(required);
    }
    return true;
}

}  // namespace

bool SecretStoreDpapi::SmokeTest(std::wstring* error_message) const {
    const std::wstring plain_text = L"mytools-native-dpapi-smoke";
    DATA_BLOB input = BlobFromString(plain_text);
    DATA_BLOB encrypted{};
    DATA_BLOB decrypted{};

    if (!CryptProtectData(&input, L"MyTools Native smoke test", nullptr, nullptr, nullptr, 0, &encrypted)) {
        if (error_message != nullptr) {
            *error_message = L"CryptProtectData failed.";
        }
        return false;
    }

    const BOOL unprotect_ok =
        CryptUnprotectData(&encrypted, nullptr, nullptr, nullptr, nullptr, 0, &decrypted);
    FreeBlob(&encrypted);

    if (!unprotect_ok) {
        if (error_message != nullptr) {
            *error_message = L"CryptUnprotectData failed.";
        }
        return false;
    }

    const bool same = StringFromBlob(decrypted) == plain_text;
    FreeBlob(&decrypted);
    if (!same && error_message != nullptr) {
        *error_message = L"DPAPI round-trip returned unexpected data.";
    }
    return same;
}

bool SecretStoreDpapi::ProtectUtf8ToBase64(const std::string& plain_utf8,
                                           std::string* protected_base64,
                                           std::wstring* error_message) const {
    if (protected_base64 == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"ProtectUtf8ToBase64 requires an output buffer.";
        }
        return false;
    }
    if (plain_utf8.empty()) {
        if (error_message != nullptr) {
            *error_message = L"ProtectUtf8ToBase64 requires non-empty plaintext.";
        }
        return false;
    }

    DATA_BLOB input = BlobFromBytes(plain_utf8);
    DATA_BLOB encrypted{};
    if (!CryptProtectData(&input, L"MyTools Native protected UTF-8 payload", nullptr, nullptr, nullptr, 0, &encrypted)) {
        if (error_message != nullptr) {
            *error_message = L"CryptProtectData failed for UTF-8 payload.";
        }
        return false;
    }

    const bool encoded = EncodeBase64(encrypted.pbData, encrypted.cbData, protected_base64);
    FreeBlob(&encrypted);
    if (!encoded && error_message != nullptr) {
        *error_message = L"Failed to encode protected payload as base64.";
    }
    return encoded;
}

bool SecretStoreDpapi::UnprotectBase64ToUtf8(const std::string& protected_base64,
                                             std::string* plain_utf8,
                                             std::wstring* error_message) const {
    if (plain_utf8 == nullptr) {
        if (error_message != nullptr) {
            *error_message = L"UnprotectBase64ToUtf8 requires an output buffer.";
        }
        return false;
    }

    std::vector<BYTE> encrypted;
    if (!DecodeBase64(protected_base64, &encrypted)) {
        if (error_message != nullptr) {
            *error_message = L"Protected text is not valid base64 DPAPI data.";
        }
        return false;
    }

    DATA_BLOB input{};
    input.pbData = encrypted.data();
    input.cbData = static_cast<DWORD>(encrypted.size());

    DATA_BLOB decrypted{};
    const BOOL ok = CryptUnprotectData(&input, nullptr, nullptr, nullptr, nullptr, 0, &decrypted);
    SecureZeroMemory(encrypted.data(), encrypted.size());

    if (!ok) {
        if (error_message != nullptr) {
            *error_message = L"CryptUnprotectData failed for the current Windows user.";
        }
        return false;
    }

    plain_utf8->assign(reinterpret_cast<const char*>(decrypted.pbData),
                       reinterpret_cast<const char*>(decrypted.pbData) + decrypted.cbData);
    FreeBlob(&decrypted);
    return true;
}

}  // namespace mytools
