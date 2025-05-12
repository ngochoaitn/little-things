using System;

namespace ReadMailOutlook.Models
{
    internal enum ProxyType : int
    {
        NoneProxy = 0,
        HttpProxy,
        Socks5
    }

    internal class ProxyInfo
    {
        public ProxyType Type { get; set; }
        public string Host { get; set; }
        public int Port { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }

        /// <summary>
        /// + Http Proxy: IP:Port or IP:Port:User:Pass
        /// + Socks 5   : socks5://IP:Port or socks5://IP:Port:User:Pass
        /// </summary>
        /// <param name="proxyString"></param>
        /// <returns></returns>
        public ProxyInfo(string proxyString)
        {
            string proxyRawString = proxyString;

            if (string.IsNullOrEmpty(proxyRawString) || !CanParse(proxyRawString))
            {
                Type = ProxyType.NoneProxy;
                return;
            }

            this.Type = ProxyType.HttpProxy;

            string prefix = "socks5://";
            if (proxyRawString.IndexOf(prefix) == 0)
            {
                this.Type = ProxyType.Socks5;
                proxyRawString = proxyRawString.Replace(prefix, "");
            }
            
            string[] spliter = proxyRawString.Split(':');
            if (spliter.Length == 2)
            {
                this.Host = spliter[0];
                int.TryParse(spliter[1], out int port);
                this.Port = port;
            }
            else if (spliter.Length == 4)
            {
                this.Host = spliter[0];
                int.TryParse(spliter[1], out int port);
                this.Port = port;
                this.UserName = spliter[2];
                this.Password = spliter[3];
            }
        }

        #region Helpers
        private bool CanParse(string proxyRawString)
        {
            if (string.IsNullOrEmpty(proxyRawString))
                return true;

            string proxy = proxyRawString.Replace("socks5://", "");

            string[] spliter = proxy.Split(':');
            return (spliter.Length == 2 || spliter.Length == 4);
        }
        #endregion 
    }
}
