# FRP 服务器连接测试报告

**测试时间**: 2026-05-31 10:54 AM (UTC+8)  
**测试目标**: 120.26.50.234:7000 (frps 服务器)  
**测试人员**: MyTools FRP 模块测试工程师

---

## 一、测试环境

### 1.1 客户端环境
- **操作系统**: Windows 10/11 (win32 10.0.26200)
- **SSH Key**: `C:\Users\QianGua\.ssh\oa_wms_server_ed25519` (Ed25519)
- **FRP 客户端**: `frpc.exe` (13.49 MB) 已嵌入 `NativeBinaries\frpc.exe.gz`

### 1.2 服务器环境
- **服务器**: 阿里云 ECS (oa-wms-server)
- **操作系统**: Ubuntu/Linux (systemd)
- **frps 版本**: v0.50.0
- **frps 配置路径**: `/usr/local/bin/frps -c /etc/frps.ini`

---

## 二、服务器端状态检查

### 2.1 SSH 连接测试

| 项目 | 结果 |
|------|------|
| 连接方式 | Ed25519 SSH Key |
| SSH Key 文件 | `oa_wms_server_ed25519` |
| 连接超时 | 10 秒 |
| **连接状态** | ✅ **成功** |

### 2.2 frps 服务状态

```
● frps.service - frps
     Loaded: loaded (/etc/systemd/system/frps.service; enabled; preset: enabled)
     Active: active (running) since Sat 2026-05-30 17:04:07 CST; 17h ago
   Main PID: 884 (frps)
      Tasks: 5 (limit: 8650)
     Memory: 21.7M (peak: 24.5M)
```

| 项目 | 结果 |
|------|------|
| **服务状态** | ✅ **运行中** (已运行 17 小时) |
| 启动时间 | 2026-05-30 17:04:07 CST |

### 2.3 frps 端口监听状态

```
tcp6       0      0 :::7000                 :::*                    LISTEN      884/frps
```

| 项目 | 结果 |
|------|------|
| **监听端口** | ✅ **7000** 正在监听 |
| 协议 | TCPv6 (兼容 IPv4) |
| 绑定地址 | `:::` (所有接口) |
| 进程 PID | 884 |

### 2.4 frps 配置文件

```ini
[common]
bind_port = 7000
token = 53a37334cb4c82fd832609c298f488958979d857373808e3
log_file = /var/log/frps.log
log_level = info
log_max_days = 7
```

| 项目 | 值 |
|------|---|
| 服务器地址 | 120.26.50.234 |
| 控制端口 | 7000 |
| Token | `53a37334cb4c82fd832609c298f488958979d857373808e3` |

---

## 三、网络连通性测试

### 3.1 端口可达性测试

| 目标端口 | 服务 | TcpTestSucceeded |
|----------|------|-------------------|
| **7000** | frps 控制端口 | ❌ **失败** (TCP connect failed) |
| 80 | HTTP/Nginx | ✅ **成功** |

**结论**: 从本地网络无法连接到服务器的 7000 端口，但 80 端口可达。

---

## 四、防火墙检查

### 4.1 UFW 防火墙状态

```
Status: active
Logging: on (low)
Default: deny (incoming), allow (outgoing), disabled (routed)
```

**已放行的端口**:

| 端口 | 协议 | 用途 | 状态 |
|------|------|------|------|
| 22 | TCP | SSH | ✅ 已放行 |
| 80 | TCP | HTTP (nginx) | ✅ 已放行 |
| 443 | TCP | HTTPS | ✅ 已放行 |
| 5173 | TCP | OA WMS dev frontend | ✅ 已放行 |
| 8001 | TCP | OA WMS dev backend API | ✅ 已放行 |
| **7000** | TCP | **frps** | ❌ **未放行** |

### 4.2 iptables 检查

```
Chain INPUT (policy DROP)
target     prot opt source               destination         
ufw-before-logging-input  0    --  0.0.0.0/0            0.0.0.0/0           
ufw-after-input  0    --  0.0.0.0/0            0.0.0.0/0           
ufw-reject-input  0    --  0.0.0.0/0            0.0.0.0/0           
```

**默认策略**: DROP (拒绝所有未匹配的入站流量)

---

## 五、问题根因分析

### 5.1 根本原因

**❌ UFW 防火墙未放行 7000 端口**

frps 服务虽然正在监听 7000 端口，但由于 UFW 的默认策略是 `DROP`，且没有显式放行 7000 端口，导致外部流量被防火墙拦截。

### 5.2 可能的安全组问题

阿里云 ECS 安全组可能也存在相同问题，需要同时检查安全组规则是否放行了 7000 端口。

---

## 六、客户端配置要求

根据 `FrpService.cs` 和 `FrpViewModel.cs` 分析：

### 6.1 客户端必需配置

| 参数 | 默认值 | 来源 |
|------|--------|------|
| 服务器地址 | 120.26.50.234 | `FrpDefaults.DefaultServerAddress` |
| 服务器端口 | 7000 | `FrpDefaults.DefaultServerPort` |
| Token | 需用户填写 | 需与 `/etc/frps.ini` 中的 token 匹配 |
| 客户端类型 | TCP | 当前版本仅支持 TCP 隧道 |

### 6.2 frpc.ini 格式示例

```ini
[common]
server_addr = 120.26.50.234
server_port = 7000
token = 53a37334cb4c82fd832609c298f488958979d857373808e3

[mytools_pc_client_3389_33890]
type = tcp
local_ip = 127.0.0.1
local_port = 3389
remote_port = 33890
```

---

## 七、解决方案

### 7.1 服务器端修复（需要 SSH 访问）

#### 方案 A: 放行 UFW 7000 端口

```bash
# 放行 7000 端口
sudo ufw allow 7000/tcp

# 重载防火墙
sudo ufw reload

# 验证
sudo ufw status | grep 7000
```

#### 方案 B: 使用 iptables 直接放行

```bash
# 放行 7000 端口
sudo iptables -A INPUT -p tcp --dport 7000 -j ACCEPT

# 保存规则 (Ubuntu)
sudo netfilter-persistent save
```

### 7.2 阿里云安全组检查

在阿里云控制台检查安全组规则，确保以下端口已放行：

| 方向 | 协议 | 端口范围 | 来源 |
|------|------|----------|------|
| 入方向 | TCP | 7000 | 0.0.0.0/0 |

### 7.3 修复后验证步骤

1. **服务器端验证**:
   ```bash
   ss -tlnp | grep 7000
   # 应该显示: LISTEN 0 4096 *:7000 *:* 
   ```

2. **客户端验证**:
   ```powershell
   Test-NetConnection -ComputerName 120.26.50.234 -Port 7000
   # 应该显示: TcpTestSucceeded = True
   ```

3. **程序内测试**:
   - 启动 MyTools
   - 填写服务器地址: `120.26.50.234`
   - 填写服务器端口: `7000`
   - 填写 Token: `53a37334cb4c82fd832609c298f488958979d857373808e3`
   - 添加一条隧道规则（如 RDP: 3389 → 33890）
   - 点击「启动隧道」
   - 检查日志中是否出现 `login to server success`

---

## 八、测试结论

### 8.1 服务器状态汇总

| 检查项 | 状态 | 说明 |
|--------|------|------|
| SSH 连接 | ✅ 正常 | 使用 Ed25519 Key 连接成功 |
| frps 服务 | ✅ 运行中 | 已运行 17 小时，无异常 |
| 7000 端口监听 | ✅ 正常 | frps 正确绑定 7000 |
| UFW 防火墙 | ❌ **异常** | **7000 端口未放行** |
| 客户端连通性 | ❌ 失败 | TCP 连接被拒绝 |

### 8.2 最终判定

**❌ FRP 客户端无法连接服务器**

**原因**: UFW 防火墙默认策略为 DROP，且未放行 7000 端口，导致外部客户端无法连接到 frps 控制端口。

### 8.3 修复优先级

| 优先级 | 任务 | 负责方 |
|--------|------|--------|
| P0 | 在服务器上放行 7000/tcp 端口 | 服务器管理员 |
| P1 | 检查阿里云安全组是否放行 7000 | 服务器管理员 |
| P2 | 验证客户端连通性 | 测试工程师 |
| P3 | 使用 MyTools 程序内测试 | 测试工程师 |

---

## 九、附录

### 9.1 frps 配置参考

```
配置文件路径: /etc/frps.ini
日志文件: /var/log/frps.log
日志级别: info
Token: 53a37334cb4c82fd832609c298f488958979d857373808e3
```

### 9.2 frpc 客户端版本

| 文件 | 大小 | 说明 |
|------|------|------|
| `NativeBinaries/frpc.exe` | 13.49 MB | 未压缩版本 |
| `NativeBinaries/frpc.exe.gz` | 5.22 MB | 压缩版本（程序内使用） |

### 9.3 端口预设参考（MyTools 内置）

| 预设名称 | 本地端口 | 远程端口 |
|----------|----------|----------|
| 远程桌面 (RDP) | 3389 | 33890 |
| 网页 HTTP 80 | 80 | 8081 |
| 网页开发 8080 | 8080 | 8082 |
| SSH 远程 | 22 | 2222 |
| MySQL 数据库 | 3306 | 3307 |

> **注意**: 使用远程桌面预设时，远程端口 33890 也需要在防火墙和安全组中放行。

---

**报告生成时间**: 2026-05-31 10:54 AM  
**测试工具**: MyTools FRP 模块 / SSH / PowerShell Test-NetConnection
