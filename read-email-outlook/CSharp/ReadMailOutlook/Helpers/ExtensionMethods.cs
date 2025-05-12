using MailKit.Net.Imap;
using ReadMailOutlook.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ReadMailOutlook.Helpers
{
    internal static class ExtensionMethods
    {
        public static void SetProxy(this ImapClient client, ProxyInfo proxy)
        {
            if (proxy == null)
                return;

            if (proxy.Type == ProxyType.HttpProxy)
            {
                if (!string.IsNullOrEmpty(proxy.UserName))
                    client.ProxyClient = new MailKit.Net.Proxy.HttpProxyClient(proxy.Host, proxy.Port, new NetworkCredential(proxy.UserName, proxy.Password));
                else
                    client.ProxyClient = new MailKit.Net.Proxy.HttpProxyClient(proxy.Host, proxy.Port);
            }
            else if (proxy.Type == ProxyType.Socks5)
            {
                if (!string.IsNullOrEmpty(proxy.UserName))
                    client.ProxyClient = new MailKit.Net.Proxy.Socks5Client(proxy.Host, proxy.Port, new NetworkCredential(proxy.UserName, proxy.Password));
                else
                    client.ProxyClient = new MailKit.Net.Proxy.Socks5Client(proxy.Host, proxy.Port);
            }
        }
    }
}
