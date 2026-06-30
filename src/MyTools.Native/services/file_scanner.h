#pragma once

#include <atomic>
#include <string>
#include <vector>

namespace mytools {

class ScanCancellation {
public:
    void Cancel() noexcept { cancelled_.store(true); }
    bool IsCancellationRequested() const noexcept { return cancelled_.load(); }

private:
    std::atomic_bool cancelled_ = false;
};

struct FileScanOptions {
    std::wstring root_directory;
    bool recursive = true;
    size_t max_results = 10000;
};

struct FileScanResult {
    std::vector<std::wstring> files;
    bool cancelled = false;
    std::wstring error_message;
};

class FileScanner {
public:
    FileScanResult ScanFiles(const FileScanOptions& options, const ScanCancellation& cancellation) const;
};

}  // namespace mytools
