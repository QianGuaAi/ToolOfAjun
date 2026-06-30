#include "services/network_service.h"

#include <sstream>

#include <winsock2.h>
#include <ws2tcpip.h>

namespace mytools {
namespace {

class WinSockSession {
public:
    WinSockSession() {
        WSADATA data{};
        startup_error_ = WSAStartup(MAKEWORD(2, 2), &data);
    }

    ~WinSockSession() {
        if (startup_error_ == 0) {
            WSACleanup();
        }
    }

    bool ok() const noexcept { return startup_error_ == 0; }
    int startup_error() const noexcept { return startup_error_; }

private:
    int startup_error_ = 0;
};

std::wstring ToPortText(uint16_t port) {
    std::wstringstream stream;
    stream << port;
    return stream.str();
}

TcpProbeResult Failure(int error, const std::wstring& message) {
    TcpProbeResult result;
    result.connected = false;
    result.winsock_error = error;
    result.message = message;
    return result;
}

}  // namespace

TcpProbeResult NetworkService::ProbeTcp(const std::wstring& host,
                                        uint16_t port,
                                        uint32_t timeout_ms) const {
    if (host.empty() || port == 0) {
        return Failure(WSAEINVAL, L"TCP probe requires a host and non-zero port.");
    }

    WinSockSession session;
    if (!session.ok()) {
        return Failure(session.startup_error(), L"WSAStartup failed.");
    }

    addrinfoW hints{};
    hints.ai_family = AF_UNSPEC;
    hints.ai_socktype = SOCK_STREAM;
    hints.ai_protocol = IPPROTO_TCP;

    addrinfoW* addresses = nullptr;
    const std::wstring port_text = ToPortText(port);
    const int resolve_error = GetAddrInfoW(host.c_str(), port_text.c_str(), &hints, &addresses);
    if (resolve_error != 0) {
        return Failure(resolve_error, L"GetAddrInfoW failed.");
    }

    TcpProbeResult last_error = Failure(WSAEHOSTUNREACH, L"No address was reachable.");
    for (addrinfoW* address = addresses; address != nullptr; address = address->ai_next) {
        SOCKET socket_handle =
            socket(address->ai_family, address->ai_socktype, address->ai_protocol);
        if (socket_handle == INVALID_SOCKET) {
            last_error = Failure(WSAGetLastError(), L"socket failed.");
            continue;
        }

        u_long non_blocking = 1;
        ioctlsocket(socket_handle, FIONBIO, &non_blocking);
        const int connect_result = connect(socket_handle, address->ai_addr, static_cast<int>(address->ai_addrlen));
        if (connect_result == 0) {
            closesocket(socket_handle);
            FreeAddrInfoW(addresses);
            TcpProbeResult result;
            result.connected = true;
            result.message = L"Connected.";
            return result;
        }

        int error = WSAGetLastError();
        if (error == WSAEWOULDBLOCK || error == WSAEINPROGRESS || error == WSAEINVAL) {
            fd_set write_set;
            FD_ZERO(&write_set);
            FD_SET(socket_handle, &write_set);

            timeval timeout{};
            timeout.tv_sec = static_cast<long>(timeout_ms / 1000);
            timeout.tv_usec = static_cast<long>((timeout_ms % 1000) * 1000);
            const int selected = select(0, nullptr, &write_set, nullptr, &timeout);
            if (selected > 0 && FD_ISSET(socket_handle, &write_set)) {
                int socket_error = 0;
                int socket_error_size = sizeof(socket_error);
                getsockopt(socket_handle,
                           SOL_SOCKET,
                           SO_ERROR,
                           reinterpret_cast<char*>(&socket_error),
                           &socket_error_size);
                if (socket_error == 0) {
                    closesocket(socket_handle);
                    FreeAddrInfoW(addresses);
                    TcpProbeResult result;
                    result.connected = true;
                    result.message = L"Connected.";
                    return result;
                }
                error = socket_error;
            } else if (selected == 0) {
                error = WSAETIMEDOUT;
            } else {
                error = WSAGetLastError();
            }
        }

        closesocket(socket_handle);
        last_error = Failure(error, L"TCP connect failed.");
    }

    FreeAddrInfoW(addresses);
    return last_error;
}

}  // namespace mytools
