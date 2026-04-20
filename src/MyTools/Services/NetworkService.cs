using System;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net.Sockets;

namespace MyTools.Services
{
    public class NetworkData
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string IpAddress { get; set; }
        public string Gateway { get; set; }
        public string Dns { get; set; }
        public OperationalStatus Status { get; set; }
    }

    public static class NetworkService
    {
        public static System.Collections.Generic.List<NetworkData> GetAllNetworkDetails()
        {
            var list = new System.Collections.Generic.List<NetworkData>();
            try
            {
                var interfaces = NetworkInterface.GetAllNetworkInterfaces()
                    .Where(i => i.NetworkInterfaceType != NetworkInterfaceType.Loopback);

                foreach (var ni in interfaces)
                {
                    var props = ni.GetIPProperties();
                    var ipv4Addresses = props.UnicastAddresses
                        .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
                        .Select(a => a.Address.ToString());
                    
                    var ipv6Addresses = props.UnicastAddresses
                        .Where(a => a.Address.AddressFamily == AddressFamily.InterNetworkV6)
                        .Select(a => a.Address.ToString());

                    list.Add(new NetworkData
                    {
                        Name = ni.Name,
                        Description = ni.Description,
                        Status = ni.OperationalStatus,
                        IpAddress = string.Join("\n", ipv4Addresses.Concat(ipv6Addresses)),
                        Gateway = string.Join("\n", props.GatewayAddresses.Select(g => g.Address.ToString())),
                        Dns = string.Join("\n", props.DnsAddresses.Select(d => d.ToString()))
                    });
                }
            }
            catch { }
            return list;
        }
    }
}
