using System;
using System.Configuration;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

using HtmlAgilityPack;
using rz = RazorEngine;

namespace MercariChecker
{
    class Program
    {
        static StreamWriter sw = new StreamWriter(@"D:\tools\MercariChecker\MercariChecker.log", true, Encoding.GetEncoding("Shift_JIS"));

        static void Main(params string[] args)
        {
            if(args.Length == 0) return;

            using (sw)
            {
                Console.SetOut(sw);

                WaitCallback waitCallback = new WaitCallback(Run);
                foreach (var key in args)
                {

                    ThreadPool.QueueUserWorkItem(waitCallback, key); 
                }

                Console.ReadLine();
            }
        }

        private static void Run(object args)
        {
            string key = args.ToString();
            string url = string.Empty, category = string.Empty, user = string.Empty, searchKeyword = string.Empty, lastItemID = string.Empty, href = string.Empty, matchKeyword = string.Empty, itemID = string.Empty;
            DateTime herthCheck = DateTime.Now;
            MercariModel model;
            ItemModel item;
            HtmlNodeCollection nodes;

            var sendIDs = new StringDictionary();
            while (true)
            {
                try
                {
                    url = ConfigurationManager.AppSettings["Mercari_Url"];
                    category = ConfigurationManager.AppSettings[string.Format("{0}_Category", key)];
                    user = ConfigurationManager.AppSettings[string.Format("{0}_User", key)];
                    if (string.IsNullOrEmpty(category) == false)
                    {
                        url = string.Format("{0}category/{1}", url, category);
                    }
                    else if (string.IsNullOrEmpty(user) == false)
                    {
                        url = string.Format("{0}u/{1}", url, user);
                    }
                    else
                    {
                        searchKeyword = ConfigurationManager.AppSettings[string.Format("{0}_Search", key)];
                        url = string.Format("{0}search/", url);
                    }
                    lastItemID = ConfigurationManager.AppSettings[string.Format("{0}_LastItemID", key)];
                    model = new MercariModel();
                    itemID = string.Empty;

                    for (int i = 1; i <= 1; i++)
                    {
                        nodes = GetHtmlDocument(url, i, searchKeyword).DocumentNode.SelectNodes(@"//section[@class=""items-box""]");

                        if (nodes == null) continue;

                        foreach (HtmlNode node in nodes)
                        {
                            href = node.SelectSingleNode(@".//a").Attributes.Single(tag => tag.Name == "href").Value;
                            itemID = href.Substring(href.LastIndexOf("m"), 10);

                            if (node.SelectSingleNode(@".//div[@class=""item-sold-out-badge""]") == null && IsInKeywords(key, href, out matchKeyword))
                            {
                                if (model.Items.Where(e => e.DetailUrl == href).Count() == 0)
                                {
                                    item = new ItemModel();
                                    item.DetailUrl = href;
                                    item.Name = node.SelectSingleNode(@".//h3[@class=""items-box-name font-2""]").InnerText;
                                    item.ImgUrl = node.SelectSingleNode(@".//img").Attributes.Single(tag => tag.Name == "data-src").Value;
                                    item.Price = node.SelectSingleNode(@".//div[@class=""items-box-price font-5""]").InnerText;
                                    item.Keyword = matchKeyword;
                                    if (sendIDs.ContainsKey(itemID) == false || sendIDs[itemID] != item.Price)
                                    {
                                        model.Items.Add(item);
                                        if (sendIDs.ContainsKey(itemID)) sendIDs[itemID] = item.Price;
                                        else sendIDs.Add(itemID, item.Price);

                                        Console.WriteLine(string.Format("{0}:{1} {2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), itemID, item.Price));
                                        Console.Out.Flush();
                                    }
                                }
                            }
                        }
                    }

                    // メール送信
                    if (model.Items.Count > 0)
                    {
                        var template = File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Mail.cshtml"));
                        var body = rz.Razor.Parse(template, model);
                        SendMail(key, body);

                        Console.WriteLine(string.Format("{0}:{1} {2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), key, model.Items.Count));
                        Console.Out.Flush();
                    }
                    else
                    {
                        TimeSpan ts = DateTime.Now - herthCheck;
                        if (ts.Minutes > 60)
                        {
                            Console.WriteLine(string.Format("Herth Check:{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
                            Console.Out.Flush();
                            herthCheck = DateTime.Now;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(string.Format("{0} Main", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
                    Console.WriteLine(ex.Message);
                    Console.WriteLine(ex.StackTrace);
                    Console.Out.Flush();
                }
                finally { sw.Flush(); }
            }
        }

        private static HtmlDocument GetHtmlDocument(string url, int page = -1, string keywords = "")
        {

            WebClient web = new WebClient();
            HtmlDocument doc = new HtmlDocument();
            string html;
            try
            {
                web.Encoding = Encoding.UTF8;
                if (string.IsNullOrEmpty(keywords) == false)
                {
                    NameValueCollection nvc = new NameValueCollection();
                    nvc.Add("keyword", HttpUtility.UrlEncode(keywords, Encoding.UTF8));
                    nvc.Add("page", page.ToString());
                    web.QueryString = nvc;
                }
                html = web.DownloadString(url);
                doc.LoadHtml(html);
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0} GetHtmlDocument", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.Out.Flush(); 
            }

            if (web != null) web.Dispose();

            return doc;
        }

        private static bool IsInKeywords(string key, string url, out string matchKeyword)
        {
            matchKeyword = string.Empty;
            HtmlDocument doc = GetHtmlDocument(url);

            // 該当ページなし
            if (doc.DocumentNode.SelectSingleNode(@"//div[@class=""item-name item-base""]") == null) return false;
            // 該当商品なし
            if (doc.DocumentNode.SelectSingleNode(@"//h1[@class=""errors-box""]") != null) return false;

            string name = doc.DocumentNode.SelectSingleNode(@"//div[@class=""item-name item-base""]").InnerText;
            string description = doc.DocumentNode.SelectSingleNode(@"//div[@class=""item-description f14""]").InnerText;

            // NG
            foreach (string keyword in ConfigurationManager.AppSettings[string.Format("{0}_Match_Keywords_NG", key)].Split(','))
            {
                if ((string.IsNullOrEmpty(name) == false && name.ToLower().IndexOf(keyword.ToLower()) > -1) ||
                    (string.IsNullOrEmpty(description) == false && description.ToLower().IndexOf(keyword.ToLower()) > -1))
                {
                    return false;
                }
            }

            // OK
            int match;
            string keywords = ConfigurationManager.AppSettings[string.Format("{0}_Match_Keywords_OK", key)];
            if (string.IsNullOrEmpty(keywords)) return true;
            foreach (string keyword in keywords.Split(','))
            {
                match = keyword.Split(' ').Length;
                foreach (string word in keyword.Split(' '))
                {
                    if ((string.IsNullOrEmpty(name) == false && name.ToLower().IndexOf(word.ToLower()) > -1) ||
                        (string.IsNullOrEmpty(description) == false && description.ToLower().IndexOf(word.ToLower()) > -1))
                    {
                        match -= 1;
                    }
                }
                if (match == 0) { 
                    matchKeyword = keyword;
                    return true;
                }
            }

            return false;
        }

        private static void SetLastItemID(string key, string itemID)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            try
            {
                config.AppSettings.Settings[string.Format("{0}_LastItemID", key)].Value = itemID;
                config.Save(ConfigurationSaveMode.Modified); 
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0} SetLastItemID", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")));
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.Out.Flush();
            }
            finally
            {
                ConfigurationManager.RefreshSection("appSettings");
            }
        }

        private static void SendMail(string key, string body)
        {
            MailMessage mm = new MailMessage();
            //SMTPサーバーを指定する
            SmtpClient smtp = new SmtpClient(); 
            try
            {
                //本文の文字コードを指定する
                mm.BodyEncoding = Encoding.UTF8;
                //あて先
                foreach (string to in ConfigurationManager.AppSettings[string.Format("{0}_Mail_To", key)].Split(',')) {
                    mm.To.Add(new MailAddress(to));
                }
                //件名
                mm.Subject = string.Format("[mercari] {0}出品最新情報", key);
                //本文
                //HTMLメールとする
                mm.IsBodyHtml = true;
                mm.Body = body;
                //送信する
                smtp.Send(mm);
            }
            finally
            { 
                //後始末
                mm.Dispose();
                smtp.Dispose();
            }
        }
    }
}
