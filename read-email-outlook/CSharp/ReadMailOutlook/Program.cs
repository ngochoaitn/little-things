using ReadMailOutlook.Helpers;
using System;
using System.Text;

namespace ReadMailOutlook
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            // 1 => email|pass|refreshToken|clientId
            // 2 => email|refreshToken|clientId
            // 3 => email|pass                       // This option may trigger account verification
            string data = "email@outlook.com.vn|password|refresh_token|client_id";
            OutlookHelper outlookHelper = new OutlookHelper(data);

            //string accessToken = outlookHelper.GetAccessToken();
            //string token = outlookHelper.GetOAuth2Token();
            var mails = outlookHelper.GetEmails();
            Console.WriteLine($"Get {mails.Count} emails");
            Console.WriteLine($"===========");

            for (int i = 0; i < mails.Count; i++)
            {
                var mail = mails[i];
                Console.WriteLine($"Mail {i+1}: {mail.Subject.Trim()}");
            }

            Console.WriteLine($"===========");
            Console.WriteLine("Press Enter to exit");
            Console.ReadLine();
        }
    }
}
