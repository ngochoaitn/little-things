using MimeKit;

namespace ReadMailOutlook.Models
{
    internal class MailData
    {
        public MailData() { }
        public MailData(MimeMessage message)
        {
            From = message.From[0].ToString();
            Subject = message.Subject;
            Html = message.HtmlBody ?? message.TextBody;
            DateTime = message.Date.ToString("yyyy-MM-dd HH:mm:ss");
            try
            {
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(Html);
                Content = doc.DocumentNode.InnerText;
            }
            catch { }
        }
        public string From { get; set; }
        public string Content { get; set; }
        public string Html { get; set; }
        public string Subject { get; set; }
        public string DateTime { get; set; }
    }
}
