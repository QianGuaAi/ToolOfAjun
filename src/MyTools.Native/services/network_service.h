#pragma once

#include <cstdint>
#include <string>

namespace mytools {

struct TcpProbeResult {
    bool connected = false;
    int winsock_error = 0;
    std::wstring message;
};

class NetworkService {
public:
    TcpProbeResult ProbeTcp(const std::wstring& host, uint16_t port, uint32_t timeout_ms) const;
};

}  // namespace mytools
