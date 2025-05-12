using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using Newtonsoft.Json.Linq;
using ReadMailOutlook.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace ReadMailOutlook.Helpers
{
    /// <summary>
    /// Github: https://github.com/ngochoaitn/little-things
    /// </summary>
    class OutlookHelper
    {
        private string _clientId;
        private readonly string redirectUri = "https://localhost";
        private readonly string baseUrl = "https://login.live.com";
        private readonly string tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";

        Leaf.xNet.CookieStorage cookieStorage = new Leaf.xNet.CookieStorage();
        private string _accessToken = "";
        private string _email = "";
        private string _password = "";
        private string _refreshToken = "";

        ProxyInfo _proxy;

        /// <summary>
        /// [1] => email|pass|refresh_token|client_id or [2] => email|pass|client_id or [3] => email|pass
        /// </summary>
        /// <param name="data"></param>
        /// <param name="proxy"></param>
        public OutlookHelper(string data, ProxyInfo proxy = null)
        {
            var datas = data.Split('|');
            if (datas.Length == 2) // email|pass
            {
                _email = datas[0];
                _password = datas[1];
                _clientId = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"; // Mozilla Thunderbird
            }
            else if (datas.Length == 3) // email|refreshToken|clientId
            {
                _email = datas[0];
                _refreshToken = datas[1];
                _clientId = datas[2];
            }
            else if (datas.Length == 4) // email|pass|refreshToken|clientId
            {
                _email = datas[0];
                _password = datas[1];
                _refreshToken = datas[2];
                _clientId = datas[3];
            }

            _proxy = proxy;
            if (_proxy == null)
                _proxy = new ProxyInfo("");
        }

        [Obsolete("This option may trigger account verification")]
        public OutlookHelper(string email, string password, ProxyInfo proxy = null)
        {
            _email = email;
            _password = password;
            _proxy = proxy;
            _clientId = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"; // Mozilla Thunderbird
            if (_proxy == null)
                _proxy = new ProxyInfo("");
        }

        public OutlookHelper(string email, string refreshToken, string clientId, ProxyInfo profileProxy = null)
        {
            _email = email;
            _refreshToken = refreshToken;
            _clientId = clientId;
            _proxy = profileProxy;
            if (_proxy == null)
                _proxy = new ProxyInfo("");
        }

        public OutlookHelper(string email, string password, string refreshToken, string clientId, ProxyInfo profileProxy = null)
        {
            _email = email;
            _password = password;
            _refreshToken = refreshToken;
            _clientId = clientId;
            _proxy = profileProxy;
            if (_proxy == null)
                _proxy = new ProxyInfo("");
        }

        /// <summary>
        /// Default read: inbox, spam
        /// </summary>
        /// <param name="senderKeyword">Filter by sender keyword</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public List<MailData> GetEmails(string senderKeyword = "@", int maxGetEmail = 10)
        {
            List<MailData> result = new List<MailData>();

            if (!ValidAccessToken())
                throw new Exception("OAuth2 outlook error");

            if (string.IsNullOrEmpty(senderKeyword))
                senderKeyword = "@";

            try
            {
                using (ImapClient client = new ImapClient())
                {
                    client.SetProxy(_proxy);

                    client.Connect("outlook.office365.com", 993, SecureSocketOptions.SslOnConnect);
                    var oauth2 = new SaslMechanismOAuth2(_email, _accessToken);
                    client.Authenticate(oauth2);
                    List<IMailFolder> lstMailBoxReads = GetMailBoxFolderToRead(client);

                    foreach (var mailbox in lstMailBoxReads)
                    {
                        mailbox.Open(FolderAccess.ReadOnly);

                        var searchQuery = SearchQuery.FromContains(senderKeyword);//.And(SearchQuery.New);
                        var uids = mailbox.Search(searchQuery);

                        for (int i = uids.Count - 1; i >= 0; i--)
                        {
                            var message = mailbox.GetMessage(uids[i]);

                            MailData data = new MailData(message);
                            result.Add(data);

                            if (result.Count > maxGetEmail)
                                break;
                        }
                    }
                }
            }
            catch
            {
                if (!string.IsNullOrEmpty(_password))
                {
                    foreach (string mailBox in new string[] { "inbox", "junkemail" })
                    {
                        var temp = this.GetEmailsByRequest(mailBox, maxGetEmail, senderKeyword);
                        result.AddRange(temp);
                    }
                }
            }

            return result.OrderByDescending(p => p.DateTime).ToList();
        }

        /// <summary>
        /// Warn: This option may trigger account verification
        /// </summary>
        /// <param name="folder">inbox, junkemail,...</param>
        /// <returns></returns>
        [Obsolete("This option may trigger account verification")]
        List<MailData> GetEmailsByRequest(string folder, int maxGetMail = 10, string senderKeyword = "@")
        {
            if (string.IsNullOrEmpty(senderKeyword))
                senderKeyword = "@";

            List<MailData> res = new List<MailData>();
            if (!ValidAccessToken())
                return res;
            using (var client = CreatexNetRequest())
            {
                client.Authorization = $"Bearer {_accessToken}";
                string url = $"https://graph.microsoft.com/v1.0/users/{_email}/mailFolders/{folder}/messages?$top={maxGetMail}&$filter=contains(from/emailAddress/address, '{senderKeyword}')";

                var response = client.Get(url);
                string responseString = response.ToString();

                if (response.StatusCode == Leaf.xNet.HttpStatusCode.OK)
                {
                    JObject json = JObject.Parse(responseString);
                    var messages = json["value"];

                    foreach (var message in messages)
                    {
                        MailData data = new MailData();
                        data.Subject = message["subject"]?.ToString() ?? "";
                        if (message["from"] != null && message["from"]["emailAddress"] != null)
                            data.From = message["from"]["emailAddress"]["address"]?.ToString() ?? "";

                        data.DateTime = message["receivedDateTime"].ToString();
                        if (DateTime.TryParse(data.DateTime, out DateTime temp))
                            data.DateTime = temp.ToString("yyyy-MM-dd HH:mm:ss");

                        if (message["body"] != null)
                        {
                            data.Html = message["body"]["content"]?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(data.Html))
                            {
                                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                                doc.LoadHtml(data.Html);
                                data.Content = doc.DocumentNode.InnerText;
                            }
                        }
                        res.Add(data);
                    }
                }
                else
                {
                    Debug.WriteLine("Lỗi lấy email: " + responseString);
                }
            }
            return res.OrderByDescending(p => p.DateTime).ToList();
        }

        /// <summary>
        /// Warn: This option may trigger account verification
        /// Return access token, refresh token.
        /// Source python: https://taphoammo.net/post/detail/chia-se-cach-lay-refreshtoken-tu-tai-khoan-hotmail-outlook_976752
        /// </summary>
        /// <param name="email"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        [Obsolete("This option may trigger account verification")]
        public string GetOAuth2Token(string email, string password)
        {
            string authUrl = $"{baseUrl}/oauth20_authorize.srf";
            var queryParams = new Dictionary<string, string>
            {
                { "response_type", "code" },
                { "client_id", _clientId },
                { "redirect_uri", redirectUri },
                { "scope", "offline_access Mail.ReadWrite" },
                { "login_hint", email }
            };

            var queryString = new FormUrlEncodedContent(queryParams).ReadAsStringAsync().Result;
            authUrl += "?" + queryString;

            var getHeaders = CreateGetHeaders();
            var postHeaders = CreateGetHeaders(new Dictionary<string, string> { { "content-type", "application/x-www-form-urlencoded" } });

            Leaf.xNet.HttpRequest clientGet = CreatexNetRequest();
            clientGet.ClearAllHeaders();
            foreach (var header in getHeaders)
            {
                clientGet.AddHeader(header.Key, header.Value);
            }
            var resp = clientGet.Get(authUrl);
            var respContent = resp.ToString();

            var match = Regex.Match(respContent, "https://login.live.com/ppsecure/post.srf\\?(.*?)',");
            if (!match.Success)
                return null;

            // https://login.live.com/ppsecure/post.srf?client_id=9e5f94bc-e8a4-4e73-b8be-63364c29d753&contextid=A084EDC50211513D&opid=ACFA99AA313BC656&bk=1743493951&uaid=3778115744294d3ba1df08fb9c541069&pid=15216
            string postUrl = baseUrl + "/ppsecure/post.srf?" + match.Groups[1].Value;
            string ppft = Regex.Match(respContent, "<input type=\"hidden\" name=\"PPFT\" id=\"(.*?)\" value=\"(.*?)\"").Groups[2].Value;

            var loginData = new Dictionary<string, string>
                {
                    { "ps", "2" },
                    { "PPFT", ppft },
                    { "PPSX", "Passp" },
                    { "NewUser", "1" },
                    { "login", email },
                    { "loginfmt", email },
                    { "passwd", password },
                    { "type", "11" },
                    { "LoginOptions", "1" },
                    { "i13", "1" },
                    { "CookieDisclosure", "0" },
                    { "IsFidoSupported", "1" }
                };
            Leaf.xNet.HttpRequest clientPost = CreatexNetRequest();
            clientGet.ClearAllHeaders();
            foreach (var header in getHeaders)
            {
                clientGet.AddHeader(header.Key, header.Value);
            }
            Leaf.xNet.FormUrlEncodedContent encodedLoginData = new Leaf.xNet.FormUrlEncodedContent(loginData);
            var loginResp = clientPost.Post(postUrl, encodedLoginData);
            string redirectUrl = loginResp.Location;

            if (string.IsNullOrEmpty(redirectUrl))
            {
#if DEBUG
                string bodyDebug = resp.ToString();
#endif
                match = Regex.Match(loginResp.ToString(), "id=\"fmHF\" action=\"(.*?)\"");
                if (!match.Success)
                    return null;

                postUrl = match.Groups[1].Value;
                if (postUrl.Contains("Update?mkt="))
                {
                    redirectUrl = HandleConsentPage(postUrl, loginResp.ToString());
                }
                else if (postUrl.Contains("confirm?mkt="))
                {
                    throw new Exception("TODO: Handle confirm?mkt= - Input secret code");
                }
                else if (postUrl.Contains("Add?mkt="))
                {
                    throw new Exception("TODO: Handle Add?mkt= - Input recovery email");
                }
            }

            if (!string.IsNullOrEmpty(redirectUrl))
            {
                string code = redirectUrl.Split('=')[1];

                var tokenData = new Dictionary<string, string>
                {
                    { "code", code },
                    { "client_id", _clientId },
                    { "redirect_uri", redirectUri },
                    { "grant_type", "authorization_code" }
                };

                var tokenResponse = clientPost.Post(tokenUrl, new Leaf.xNet.FormUrlEncodedContent(tokenData));
                return tokenResponse.ToString();
            }

            return null;
        }

        public string GetAccessToken(string refreshToken, string clientId)
        {
            using (Leaf.xNet.HttpRequest client = CreatexNetRequest())
            {
                var values = new Dictionary<string, string>
                {
                    { "client_id", clientId },
                    { "refresh_token", refreshToken },
                    { "grant_type", "refresh_token" },
                };

                var content = new Leaf.xNet.FormUrlEncodedContent(values);
                var response = client.Post($"https://login.microsoftonline.com/common/oauth2/v2.0/token", content);
                string responseString = response.ToString();

                if (response.StatusCode == Leaf.xNet.HttpStatusCode.OK)
                {
                    JObject json = JObject.Parse(responseString);
                    return json["access_token"]?.ToString();
                }
                else
                {
                    throw new Exception("Error get Access Token: " + responseString);
                }
            }
        }

        public string GetAccessToken()
        {
            return GetAccessToken(_refreshToken, _clientId);
        }

        #region Helper
        private Leaf.xNet.HttpRequest CreatexNetRequest()
        {
            Leaf.xNet.HttpRequest request = new Leaf.xNet.HttpRequest
            {
                AllowAutoRedirect = false,
                Cookies = cookieStorage
            };

            if (_proxy.Type == ProxyType.HttpProxy)
            {
                if (string.IsNullOrEmpty(_proxy.UserName))
                    request.Proxy = new Leaf.xNet.HttpProxyClient(_proxy.Host, _proxy.Port);
                else
                    request.Proxy = new Leaf.xNet.HttpProxyClient(_proxy.Host, _proxy.Port, _proxy.UserName, _proxy.Password);
            }
            else if (_proxy.Type == ProxyType.Socks5)
            {
                if (string.IsNullOrEmpty(_proxy.UserName))
                    request.Proxy = new Leaf.xNet.Socks5ProxyClient(_proxy.Host, _proxy.Port);
                else
                    request.Proxy = new Leaf.xNet.Socks5ProxyClient(_proxy.Host, _proxy.Port, _proxy.UserName, _proxy.Password);
            }
            return request;
        }

        private Dictionary<string, string> CreateGetHeaders(Dictionary<string, string> additionalHeaders = null)
        {
            var headers = new Dictionary<string, string>
            {
                { "accept", "*/*" },
                //{ "accept-encoding", "gzip, deflate, br" },
                { "accept-language", "en-US,en;q=0.9" },
                { "sec-ch-ua", "\"Chromium\";v=\"104\", \" Not A;Brand\";v=\"99\", \"Google Chrome\";v=\"104\""},
                { "sec-ch-ua-mobile", "?0" },
                { "sec-ch-ua-platform", "Windows" },
                { "sec-fetch-dest", "empty" },
                { "sec-fetch-mode", "cors" },
                { "sec-fetch-site", "same-origin" },
                { "user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:128.0) Gecko/20100101 Thunderbird/128.2.3" }
            };

            if (additionalHeaders != null)
            {
                foreach (var item in additionalHeaders)
                {
                    headers[item.Key] = item.Value;
                }
            }

            return headers;
        }

        private string HandleConsentPage(string postUrl, string respContent)
        {
            var postHeaders = CreateGetHeaders(new Dictionary<string, string> { { "content-type", "application/x-www-form-urlencoded" } });

            var matches = Regex.Matches(respContent, "<input type=\"hidden\" name=\"(.*?)\" id=\"(.*?)\" value=\"(.*?)\"");
            var formData = new Dictionary<string, string>();

            foreach (Match match in matches)
            {
                formData[match.Groups[1].Value] = match.Groups[3].Value;
            }

            Leaf.xNet.HttpRequest request = CreatexNetRequest();
            request.ClearAllHeaders();
            foreach (var header in postHeaders)
            {
                request.AddHeader(header.Key, header.Value);
            }

            Leaf.xNet.FormUrlEncodedContent encodedData = new Leaf.xNet.FormUrlEncodedContent(formData);
            request.Post(postUrl, encodedData);

            formData["ucaction"] = "Yes";

            encodedData = new Leaf.xNet.FormUrlEncodedContent(formData);
            var consentResp = request.Post(postUrl, encodedData);

            if (consentResp.Location != null)
            {
                //var finalResp = await client.PostAsync(consentResp.Headers.Location.ToString(), encodedData);
                var finalResp = request.Post(consentResp.Location, encodedData);
                return finalResp.Location;
            }

            return null;
        }

        private bool ValidAccessToken()
        {
            bool ValidByEmailAndPassword()
            {
                string jsonToken = GetOAuth2Token(_email, _password);
                if (string.IsNullOrEmpty(jsonToken))
                    return false;

                JObject temp = JObject.Parse(jsonToken);
                if (temp.ContainsKey("access_token"))
                {
                    _accessToken = temp["access_token"].ToString();
                    return !string.IsNullOrEmpty(_accessToken);
                }
                return false;
            }
            if (string.IsNullOrEmpty(_accessToken))
            {
                if (!string.IsNullOrEmpty(_refreshToken))
                {
                    try
                    {
                        _accessToken = GetAccessToken(_refreshToken, _clientId);
                        return !string.IsNullOrEmpty(_accessToken);
                    }
                    catch
                    {
                        if (!string.IsNullOrEmpty(_password))
                        {
                            cookieStorage = new Leaf.xNet.CookieStorage();
                            bool checkUserClientId = ValidByEmailAndPassword();
                            if (checkUserClientId)
                                return checkUserClientId;
                            cookieStorage = new Leaf.xNet.CookieStorage();
                            _clientId = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"; // DÙng client id của Mozilla Thunderbird
                            return ValidByEmailAndPassword();
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(_password))
                {
                    return ValidByEmailAndPassword();
                }
            }
            return true;
        }

        private List<IMailFolder> GetMailBoxFolderToRead(ImapClient client)
        {
            List<IMailFolder> lstMailBoxReads = new List<IMailFolder>() { client.Inbox };

            try { lstMailBoxReads.Add(client.GetFolder(SpecialFolder.Junk)); } catch { } // Spam folder Gmail

            if (lstMailBoxReads.Count == 1)
                try { lstMailBoxReads.Add(client.GetFolder("[Gmail]/Spam")); } catch { } // Spam folder Gmail

            if (lstMailBoxReads.Count == 1)
                try { lstMailBoxReads.Add(client.GetFolder("Junk")); } catch { } // Spam folder Outlook

            if (lstMailBoxReads.Count == 1)
                try { lstMailBoxReads.Add(client.GetFolder("Spam")); } catch { } // Spam folder other mail server

            // Try to find spam folder
            if (lstMailBoxReads.Count == 1)
            {
                try
                {
                    var allMailFolders = client.GetFolders(new FolderNamespace('/', ""));
                    foreach (var folder in allMailFolders)
                    {
                        if (folder.FullName.ToLower().Contains("spam")
                            || folder.FullName.ToLower().Contains("junk"))
                        {
                            lstMailBoxReads.Add(folder);
                            break;
                        }
                    }
                }
                catch { }
            }
            return lstMailBoxReads;

        }

        #endregion
    }
}