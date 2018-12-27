using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using RestSharp;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using OpenQA.Selenium;
using SeleniumExtentions;
using System.Net;
using System.Threading.Tasks;
using System.Web;
using System.Management.Automation;

namespace ReportGenerators
{
    
    public static class StringExtentions
    {
        public static string[] CleanSplit(this string ToSplit, string seperator)
        {
            if(ToSplit == null)
            {
                return null;
            }
            return ToSplit.Split(new[] { seperator }, StringSplitOptions.RemoveEmptyEntries)
                          .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                          .ToArray();
        }
        public static string[] CleanSplit(this string ToSplit, char seperator)
        {
            if (ToSplit == null)
            {
                return null;
            }
            return ToSplit.Split(seperator)
                          .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                          .ToArray();
        }
        public static string FirstCharToUpper(this string input)
        {
            switch (input)
            {
                case null: throw new ArgumentNullException(nameof(input));
                case "": throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input));
                default: return input.First().ToString().ToUpper() + input.Substring(1);
            }
        }
    }
    //Name space for all classes needed for the ReportGenerators
    public class DataToParse
    {
        public DataToParse(string location, string page_body)
        {
            this.Location = location;
            Doc = new HtmlDocument();
            Doc.LoadHtml(page_body);
        }
        public string Location;
        public HtmlDocument Doc;
    }
   
    public class YoutubeData
    {
        public class ItemData
        {
            public class ContentDetails
            {
                public string duration { get; set; }
                public bool caption { get; set; }
                public bool licensedContent { get; set; }
            }
            public ContentDetails contentDetails { get; set; }
            public string id { get; set; }
        }
        public List<ItemData> items { get; set; }
    }

    
    public static class VideoParser
    {
        
        private static string GoogleApi = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AccessibilityTools\ReportGenerators-master\Passwords\MyGoogleApi.txt").Replace("\r\n", "");
        public static bool CheckTranscript(HtmlNode element)
        {
            if(element.NextSibling?.OuterHtml.Contains("transcript") == true
                || element.NextSibling?.NextSibling?.OuterHtml.Contains("transcript") == true
                || element.NextSibling?.NextSibling?.NextSibling?.OuterHtml.Contains("transcript") == true)
            {
                return true;
            }
            return false;
        }
        public static bool CheckTranscript(HtmlNode element, out string YesOrNo)
        {
            if (element.NextSibling.OuterHtml.Contains("transcript")
                || element.NextSibling.NextSibling.OuterHtml.Contains("transcript")
                || element.NextSibling.NextSibling.NextSibling.OuterHtml.Contains("transcript"))
            {
                YesOrNo = "Yes";
                return true;
            }
            YesOrNo = "No";
            return false;
        }
        public static TimeSpan GetYoutubeVideoLength(string video_id)
        {
            string url = $"https://www.googleapis.com/youtube/v3/videos?id={video_id}&key={GoogleApi}&part=contentDetails";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<YoutubeData>(request);
            return XmlConvert.ToTimeSpan(response.Data.items[0].contentDetails.duration);
        }
        public static TimeSpan GetBrightcoveVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://studio.brightcove.com/products/videocloud/media/videos/search/{video_id}";
            dynamic length;
            try
            {
                length = wait.Until(c =>
                {
                    var el = c.FindElement(By.CssSelector("div[class*='runtime']"));
                    if (el.Displayed)
                    {
                        return el;
                    }
                    else
                    {
                        return null;
                    }
                }).Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetBYUMediaSiteVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://byu.mediasite.com/Mediasite/Play/{video_id}";
            dynamic length = null;
            try
            {
                while("0:00" == length || "" == length || null == length)
                {
                    length = wait.Until(c =>
                    {
                        var el = c.FindElement(By.CssSelector("span[class*=\"duration\"]"));
                        if (el.Displayed)
                        {
                            return el;
                        }
                        else
                        {
                            return null;
                        }
                    }).Text;
                }
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetPanoptoVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://byu.hosted.panopto.com/Panopto/Pages/Embed.aspx?id={video_id}&amp;v=1";
            while ((string)chrome.ExecuteScript("return document.readyState") != "complete") { };
            while((int)chrome.ExecuteScript("return jQuery.active") != 0) { };
            dynamic length = null;
            try
            {
                while (chrome.FindElementsByCssSelector("[id=copyrightNoticeContainer]").FirstOrDefault().Displayed) { };
                wait.Until(c =>
                {
                    var el = c.FindElement(By.CssSelector("div[aria-label=\"Play\"]"));
                    if (el.Displayed)
                    {
                        return el;
                    }
                    else
                    {
                        return null;
                    }
                }).Click();

                length = wait.Until(c =>
                {
                    var el = c.FindElement(By.CssSelector("span[class*=\"duration\"]"));
                    if (el.Displayed)
                    {
                        return el;
                    }
                    else
                    {
                        return null;
                    }
                }).Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetAlexanderStreetVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://search.alexanderstreet.com/embed/token/{video_id}";
            dynamic length;
            try
            {
                length = wait.UntilElementIsVisible("span.fulltime").Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetAlexanderStreenLinkLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://lib.byu.edu/remoteauth/?url=https://search.alexanderstreet.com/view/work/bibliographic_entity|video_work|{video_id}";
            dynamic length;
            try
            {
                length = wait.UntilElementIsVisible("span.fulltime").Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetKanopyVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://byu.kanopy.com/embed/{video_id}";
            dynamic length;
            try
            {
                wait.UntilElementIsVisible("button.vjs-big-play-button").Click();
                length = wait.UntilElementIsVisible("div.vjs-remaining-time-display")
                                .Text
                                .Split('-')
                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                .LastOrDefault();
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetKanopyLinkLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://byu.kanopy.com/video/{video_id}";
            dynamic length;
            try
            {
                wait.UntilElementIsVisible("button.vjs-big-play-button").Click();
                length = wait.UntilElementIsVisible("div.vjs-remaining-time-display")
                                .Text
                                .Split('-')
                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                .LastOrDefault();
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetAmbroseVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            //Just realized I never implemented this function in my powershell program
            return new TimeSpan(0);
            chrome.Url = $"";
            dynamic length;
            try
            {

            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetFacebookVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://www.facebook.com/video/embed?video_id={video_id}";
            dynamic length;
            try
            {
                wait.UntilElementIsVisible("img").Click();
                length = wait.UntilElementIsVisible("div[playbackdurationtimestamp]").Text.Replace('-', '0');
            }
            catch
            {
                try
                {
                    //If that didn't work then try refreshing the page (I kept running into false negatives) and try again
                    chrome.Navigate().Refresh();
                    wait.UntilElementIsVisible("img").Click();
                    length = wait.UntilElementIsVisible("div[playbackdurationtimestamp]").Text.Replace('-', '0');
                }
                catch
                {
                    Console.WriteLine("Video not found");
                    length = "00:00";
                }
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetDailyMotionVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://www.dailymotion.com/embed/video/{video_id}";
            dynamic length;
            try
            {
                wait.UntilElementIsVisible("button[aria-label*=\"Playback\"]").Click();
                length = wait.UntilElementIsVisible("span[aria-label*=\"Duration\"]").Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
        public static TimeSpan GetVimeoVideoLength(string video_id, ChromeDriver chrome, WebDriverWait wait)
        {
            chrome.Url = $"https://player.vimeo.com/video/{video_id}";
            dynamic length;
            try
            {
                length = wait.UntilElementIsVisible("div.timecode").Text;
            }
            catch
            {
                Console.WriteLine("Video not found");
                length = "00:00";
            }
            length = "00:" + length;
            if (!TimeSpan.TryParseExact(length, @"h\:mm\:ss", null, out TimeSpan video_length))
            {
                return new TimeSpan(0);
            }
            return video_length;
        }
    }
    public abstract class RParserBase
    {
        //Base class for each of the reports
        public RParserBase() { }
        public List<PageData> Data { get; set; } = new List<PageData>();
        public DataToParse PageDocument;
        public abstract void ProcessContent(Dictionary<string, string> page_info);

    }

    public class ColorContrast
    {
        public double ratio { get; set; }
        public string AA { get; set; }
        public string AALarge { get; set; }
        public override string ToString()
        {
            var props = typeof(ColorContrast).GetProperties();
            var sb = new StringBuilder();
            foreach (var p in props)
            {
                sb.AppendLine(p.Name + ": " + p.GetValue(this, null));
            }
            return sb.ToString();
        }

    }

    public class A11yParser : RParserBase
    {
        //Class to do an accessibiltiy report
        public A11yParser() { }
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            //Function to begin processing a page and storing the data within the Data list (see RParserBase class)
            //Make sure page is not empty
            if(page_info[page_info.Keys.ElementAt(0)] == null){
                return;
            }
            //Set our current document (creates an HTML dom from the pages body)
            PageDocument = new DataToParse(page_info.Keys.ElementAt(0), page_info[page_info.Keys.ElementAt(0)]);
            //Process the elements of the page
            ProcessLinks();
            ProcessImages();
            ProcessIframes();
            ProcessTables();
            ProcessBrightcoveVideoHTML();
            ProcessHeaders();
            ProcessSemantics();
            ProcessVideoTags();
            ProcessFlash();
            ProcessColor();
        }
        private void ProcessLinks()
        {
            //Get all links within page
            var link_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//a");
            //Make sure its not null
            if(link_list == null)
            {
                return;
            }
            //Loop through all links
            foreach(var link in link_list)
            {
                
                if(link.Attributes["onclick"] != null)
                {   //Onclick links are not accessible
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.OuterHtml, "JavaScript links are not accessible", 1));
                }
                else if(link.Attributes["href"] == null)
                {   //Links should have an href
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.OuterHtml, "Empty link tag", 1));
                }
                if (link.InnerHtml.Contains("<img"))
                {   //If it is an image ignore it for now, need to check alt text
                    continue;
                }
                if(link.InnerText == null)
                {   //See if it is a link without text
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", "Invisible link with no text", "Adjust Link Text", 1));
                }else if(new Regex("^ ?here").IsMatch(link.InnerText))
                {   //If it begins with the word here probably not descriptive link text
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                }else if(new Regex("^ ?[A-Za-z\\.]+ ?$").IsMatch(link.InnerText))
                {   //If it is a single word
                    if(link_list.Where(s => s.InnerText == link.InnerText).Count() > 1)
                    {   //And if the single word is used for more then one link
                        Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                    }
                }else if(new Regex("http|www\\.|Link|Click").IsMatch(link.InnerText))
                {   //See if it is just a url
                    if(new Regex("Links to an external site").IsMatch(link.InnerText))
                    {   //This is commonly used in Canvas, we just ignore it
                        continue;
                    }
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                }
            }
        }
        private void ProcessImages()
        {
            //Get list of images
            var image_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//img");
            //Make sure it is not null
            if(image_list == null)
            {
                return;
            }
            //Loop through all images
            foreach(var image in image_list)
            {
                var alt = image.Attributes["alt"]?.Value;
                //Get the alt text
                if (alt == null)
                {   //Images should have alt tags, even if it is empty
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", image.OuterHtml, "No alt attribute", 1));
                }
                else if (new Regex("banner").IsMatch(alt))
                {   //Banners shouldn't have alt text
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if (new Regex("Placeholder").IsMatch(alt))
                {   //Placeholder probably means the alt text was forgotten to be changed
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if (new Regex("\\.jpg").IsMatch(alt))
                {   //Make sure it is not just the images file name
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("\\.png").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("http").IsMatch(alt))
                {   //It should not be a url
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("LaTeX:").IsMatch(alt))
                {   //Should not be latex (ran into this a couple of times)
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
            }
        }
        private void ProcessTables()
        {
            //Get all tables
            var table_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//table");
            //Make sure it isn't null
            if(table_list == null)
            {
                return;
            }
            //Count the tables so we can know which table on the page has the issues
            var table_num = 1;
            foreach(var table in table_list)
            {
                //Get list of headers, data cells, how many rows, and any stretched cells
                var table_headers = table.SelectNodes(".//th");
                var table_data_cells = table.SelectNodes(".//td");
                var table_rows = table.SelectNodes(".//tr");
                var stretched_cells = table.SelectNodes(".//*[@colspan]");
                //Init the issue string
                string issues = "";
                //See if there are any stretchedcells
                if(stretched_cells != null)
                {
                    issues += "\nStretched table cell(s) should be a <caption> title for the table";
                }
                //See how many rows there are, if there is 3 or more and there are no headers then flag it as needing headers
                var num_rows = table_rows.Count();
                if(num_rows >= 3)
                {
                    if(table_headers == null)
                    {
                        issues += "\nTable has no headers";
                    }
                }
                //See how many headers have scopes, should be the same number as the number of headers
                var scope_headers = table_headers?.Count(c => c.Attributes["scope"] != null);
                if(scope_headers == null || scope_headers != table_headers.Count())
                {
                    issues += "\nTable headers should have a scope attribute";
                }
                //See if any data cells have scopes when they should not
                var scope_cells = table_data_cells?.Count(c => c.Attributes["scope"] != null);
                if(scope_cells != null && scope_cells > 0)
                {
                    issues += "\nNon-header table cells should not have scope attributes";
                }
                //If any issues were found then add it to the list
                if(issues != null && issues != "")
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Table", "", $"Table number {table_num}:{issues}", "Revise table", 1));
                }
                table_num++;
            }
        }
        private void ProcessIframes()
        {
            //Get list of iframes
            var iframe_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//iframe");
            //Make sure its not null
            if(iframe_list == null)
            {
                return;
            }
            //Keep track of what iframe we are on
            var iframe_number = 1;
            foreach(var iframe in iframe_list)
            {
                //Get the source attribute, every iframe should have one
                var src = iframe.Attributes["src"].Value;
                if (iframe.Attributes["title"] == null)
                {
                    //Only real accessiblity issue we can check is if it has a title or not
                    if(new Regex("youtube").IsMatch(src))
                    {
                        //Get the youtube information
                        var uri = new Uri(src);
                        var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
                        var videoId = string.Empty;
                        if (query.AllKeys.Contains("v"))
                        {
                            videoId = query["v"];
                        }
                        else
                        {
                            videoId = uri.Segments.LastOrDefault();
                        }
                        Data.Add(new PageA11yData(PageDocument.Location, "Youtube Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("brightcove").IsMatch(src))
                    {
                        //Get brightcove info
                        var videoId = src.CleanSplit("=").LastOrDefault().CleanSplit("&").FirstOrDefault();
                        if (!src.Contains("https:"))
                        {   //Make sure it has the https on it
                            src = $"https:{src}";
                        }
                        Data.Add(new PageA11yData(PageDocument.Location, "Brightcove Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("H5P").IsMatch(src))
                    {   //H5P can just be added
                        Data.Add(new PageA11yData(PageDocument.Location, "H5P", "", "", "Needs a title", 1));
                    }else if(new Regex("byu\\.mediasite").IsMatch(src))
                    {   //Get id
                        var videoId = src.CleanSplit("/").LastOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "BYU Mediasite Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("panopto").IsMatch(src))
                    {
                        var videoId = src.CleanSplit('=').LastOrDefault().CleanSplit('&').ElementAtOrDefault(1);
                        Data.Add(new PageA11yData(PageDocument.Location, "Panopto Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("alexanderstreet").IsMatch(src))
                    {
                        var videoId = src.Split(new string[] { "token/" }, StringSplitOptions.RemoveEmptyEntries)
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "AlexanderStreen Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("kanopy").IsMatch(src))
                    {
                        var videoId = src.Split(new string[] { "embed/" }, StringSplitOptions.RemoveEmptyEntries)
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "Kanopy Video", videoId, "", "Needs a title", 1));
                    }
                    else if(new Regex("ambrosevideo").IsMatch(src))
                    {
                        var videoId = src.Split('?')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault()
                                            .Split('&')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .FirstOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "Ambrose Video", videoId, "", "NEeds a title", 1));
                    }else if(new Regex("facebook").IsMatch(src))
                    {
                        var videoId = new Regex("\\d{17}").Match(src).Value;
                        Data.Add(new PageA11yData(PageDocument.Location, "Facebook Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("dailymotion").IsMatch(src))
                    {
                        var videoId = src.Split('/')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "Facebook Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("vimeo").IsMatch(src))
                    {
                        var videoId = src.Split('/')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault()
                                            .Split('?')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .FirstOrDefault();
                        Data.Add(new PageA11yData(PageDocument.Location, "Vimeo Video", videoId, "", "Needs a title", 1));
                    }
                    else
                    {
                        Data.Add(new PageA11yData(PageDocument.Location, "Iframe", "", "", "Needs a title", 1));
                    }
                }
                if (new Regex("brightcove|byu\\.mediasite|panopto|vimeo|dailymotion|facebook|ambrosevideo|kanopy|alexanderstreet").IsMatch(src))
                {
                    if (!VideoParser.CheckTranscript(iframe))
                    {
                        Data.Add(new PageA11yData(PageDocument.Location, "Transcript", "", $"Video number {iframe_number} on page", "No transcript found", 5));
                    }
                }
                iframe_number++;
            }
        }
        private void ProcessBrightcoveVideoHTML()
        {
            var brightcove_list = PageDocument.Doc
                .DocumentNode
                ?.SelectNodes(@"//div[@id]")
                ?.Where(e => new Regex("\\d{13}").IsMatch(e.Id));
            if(brightcove_list == null)
            {
                return;
            }
            foreach (var video in brightcove_list)
            {
                if (!VideoParser.CheckTranscript(video))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Transcript", video.Attributes["id"].Value, $"No transcript found for BrightCove video with id:\n{video.Attributes["id"].Value}", "No transcript found", 5));
                }
            }
        }
        private void ProcessHeaders()
        {
            var header_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//h1 | //h2 | //h3 | //h4 | //h5 | //h6");
            if(header_list == null)
            {
                return;
            }
            foreach(var header in header_list)
            {
                if (header.Attributes["class"]?.Value?.Contains("screenreader-only") == true)
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Header", "", header.OuterHtml, "Check if header is meant to be invisible", 1));
                }
            }
        }
        private void ProcessSemantics()
        {
            var i_or_b_tag_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//i | //b");
            if(i_or_b_tag_list == null)
            {
                return;
            }else if(i_or_b_tag_list.Count() > 0)
            {
                //Flag if any i or b tags are found
                Data.Add(new PageA11yData(PageDocument.Location, "<i> or <b> tags", "", "Page contains <i> or <b> tags", "<i>/<b> tags should be <em>/<strong> tags", 1));
            }
        }
        private void ProcessVideoTags()
        {
            var videotag_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//video");
            if(videotag_list == null)
            {
                return;
            }
            foreach(var videotag in videotag_list)
            {
                var src = videotag.Attributes["src"].Value;
                var videoId = src.Split('=')
                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                    .ElementAt(1)
                                    .Split('&')
                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                    .FirstOrDefault();
                if (!VideoParser.CheckTranscript(videotag))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Inline Media Video", videoId, "Inline Media Video\n", "No transcript found", 5));
                }
            }
        }
        private void ProcessFlash()
        {
            var flash_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//object[contains(@id, \"flash\")]");
            if(flash_list == null)
            {
                return;
            }
            else if(flash_list.Count() > 0)
            {
                //Flash shouldn't be used anywhere
                Data.Add(new PageA11yData(PageDocument.Location, "Flash Element", "", $"{flash_list.Count()} embedded flash element(s) on this page", "Flash is inaccessible", 5));
            }

        }
        private void ProcessColor()
        {
            var colored_element_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//*[contains(@style, \"color\")]");
            if(colored_element_list == null)
            {
                return;
            }
            foreach(var color in colored_element_list)
            {

                System.Web.UI.CssStyleCollection style = new System.Web.UI.WebControls.Panel().Style;
                style.Value = color.Attributes["style"].Value;
                var background_color = style["background-color"];
                if(background_color == null)
                {   //Default background color is white
                    background_color = "#FFFFFF";
                }
                var foreground_color = style["color"];
                if(foreground_color == null)
                {   //Default text color is black
                    foreground_color = "#000000";
                }
                if (!background_color.Contains("#"))
                {   //If it doesn't have a # then it is a known named color, needs to be converted to hex
                    //the & 0xFFFFFF cuts ofthe A of the ARGB
                    int rgb = System.Drawing.Color.FromName(background_color.FirstCharToUpper()).ToArgb() & 0xFFFFFF;
                    background_color = string.Format("{0:x6}", rgb);
                }
                if (!foreground_color.Contains('#'))
                {   //If it doesn't have a # then it is a known named color, needs to be converted to hex
                    int rgb = System.Drawing.Color.FromName(foreground_color.FirstCharToUpper()).ToArgb() & 0xFFFFFF;
                    foreground_color = string.Format("{0:x6}", rgb);
                }
                //The API doesn't like having the #
                foreground_color = foreground_color.Replace("#", "");
                background_color = background_color.Replace("#", "");
                var restClient = new RestClient($"https://webaim.org/resources/contrastchecker/?fcolor={foreground_color}&bcolor={background_color}&api");
                var request = new RestRequest(Method.GET);
                //Will return single course object with parameters we want
                var response = restClient.Execute<ColorContrast>(request).Data;
                var text = string.Empty;
                //See if we can get the inner text so we can identify the element if there was an issue found
                if(color.InnerText != null)
                {
                    text = "\"" + HttpUtility.HtmlDecode(color.InnerText) + "\"\n";
                }
                if(response.AA != "pass")
                {   //Add it if it doesn't pass AA standards
                    Data.Add(new PageA11yData(PageDocument.Location, "Color Contrast", "", $"{text}Color: {foreground_color}\nBackgroundColor: {background_color}\n{response.ToString()}", "Does not meet AA color contrast", 1));
                }
            }
        }
    }
    
    public class MediaParser : RParserBase
    {
        private string PathToChromedriver = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AccessibilityTools\PowerShell\Modules\SeleniumTest";
        //Class to do a media report
        public MediaParser()
        {
            var chromeDriverService = ChromeDriverService.CreateDefaultService(PathToChromedriver);
            chromeDriverService.HideCommandPromptWindow = true;
            var ChromeOptions = new ChromeOptions();
            ChromeOptions.AddArguments("headless", "muteaudio");
            Chrome = new ChromeDriver(chromeDriverService, ChromeOptions);
            Wait = new WebDriverWait(Chrome, new TimeSpan(0, 0, 5));
        }
        ~MediaParser()
        {
            Chrome.Quit();
        }
        //Gen a media report
        public ChromeDriver Chrome { get; set; }
        public WebDriverWait Wait { get; set; }
        private bool LoggedIntoBrightcove = false;
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            if (!LoggedIntoBrightcove)
            {
                string BrightCoveUserName = File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AccessibilityTools\ReportGenerators-master\Passwords\MyBrightcoveUsername.txt").Replace("\n", "").Replace("\r", "");
                var posh = PowerShell.Create();
                posh.AddScript("process{$c = Get-Content \"$HOME\\Desktop\\AccessibilityTools\\ReportGenerators-master\\Passwords\\MyBrightcovePassword.txt\"; $s = $c | ConvertTo-SecureString; Write-Host (New-Object System.Management.Automation.PSCredential -ArgumentList 'asdf', $s).GetNetworkCredential().Password}"
                );
                posh.Invoke();
                var password = posh.Streams.Information[0].ToString();

                Chrome.Url = "https://signin.brightcove.com/login?redirect=https%3A%2F%2Fstudio.brightcove.com%2Fproducts%2Fvideocloud%2Fmedia";
                Wait.UntilElementIsVisible("input[name*=\"email\"]").SendKeys(BrightCoveUserName);
                Wait.UntilElementIsVisible("input[id*=\"password\"]").SendKeys(password);
                Wait.UntilElementIsVisible("button[id*=\"signin\"]").Submit();

                LoggedIntoBrightcove = true;
            }
            if (page_info[page_info.Keys.ElementAt(0)] == null)
            {
                return;
            }
            PageDocument = new DataToParse(page_info.Keys.ElementAt(0), page_info[page_info.Keys.ElementAt(0)]);

            ProcessLinks();
            ProcessIframes();
            ProcessVideoTags();
            ProcessBrightcoveVideoHTML();
        }
        private void ProcessLinks()
        {
            var link_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//a");
            if (link_list == null)
            {
                return;
            }
            foreach(var link in link_list)
            {
                if(link.Attributes["href"] == null)
                {
                    continue;
                }
                if (link.GetClasses().Contains("video_link"))
                {
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Canvas Video Link",
                                                "",
                                                "Inline Media:\nUnable to find title or video length for this type of video",
                                                link.Attributes["href"].Value,
                                                new TimeSpan(0),
                                                VideoParser.CheckTranscript(link)));
                }else if(new Regex("youtu\\.?be").IsMatch(link.Attributes["href"].Value)){
                    var split_href = link.Attributes["href"].Value
                                                                .Split(new[] { "v=" }, StringSplitOptions.RemoveEmptyEntries)
                                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                                .LastOrDefault();
                    string video_id;
                    if (split_href.Contains("t="))
                    {
                        video_id = split_href.Split('?')
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .FirstOrDefault()
                                                .Split('/')
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .LastOrDefault();
                    }else if (split_href.Contains("="))
                    {
                        video_id = split_href.Split(new[] { "v=" }, StringSplitOptions.RemoveEmptyEntries)
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .LastOrDefault()
                                                .Split('&')
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .FirstOrDefault();
                    }
                    else
                    {
                        video_id = split_href.Split('/')
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .LastOrDefault();
                    }
                    TimeSpan video_length;
                    try
                    {
                        video_id = video_id.Split('?')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .FirstOrDefault();
                        video_id = video_id.Split('/')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .LastOrDefault();
                        video_id = video_id.Split('#')
                                            .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                            .FirstOrDefault();

                        video_length = VideoParser.GetYoutubeVideoLength(video_id);
                    }
                    catch
                    {
                        Console.WriteLine("Video not found");
                        video_length = new TimeSpan(0);
                    }
                    string video_found;
                    if(video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "YouTube Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                true));
                }
                else if (link.Attributes["href"].Value.Contains("alexanderstreet"))
                {
                    string video_id = link.Attributes["href"].Value.Split('/')
                                                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                                    .LastOrDefault();
                    TimeSpan video_length = VideoParser.GetAlexanderStreenLinkLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "AlexanderStreet Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                false));
                }else if (link.Attributes["href"].Value.Contains("kanopy"))
                {
                    string video_id = link.Attributes["href"].Value.Split('/')
                                                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                                    .LastOrDefault();
                    TimeSpan video_length = VideoParser.GetAlexanderStreenLinkLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Kanopy Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                false));
                }
                else if (link.Attributes["href"].Value.Contains("byu.mediasite"))
                {
                    string video_id = link.Attributes["href"].Value.Split('/')
                                                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                                    .LastOrDefault();
                    TimeSpan video_length = VideoParser.GetAlexanderStreenLinkLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "ByuMediasite Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                false));
                }
                else if (link.Attributes["href"].Value.Contains("panopto"))
                {
                    string video_id = link.Attributes["href"].Value.Split('/')
                                                                    .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                                    .LastOrDefault();
                    TimeSpan video_length = VideoParser.GetAlexanderStreenLinkLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Panopto Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                false));
                }
                else if (link.Attributes["href"].Value.Contains("bcove"))
                {
                    Chrome.Url = link.Attributes["href"].Value;
                    Wait.UntilElementIsVisible("iframe");
                    string video_id = Chrome.Url.Split('=')
                                                .Where(s => !string.IsNullOrEmpty(s) && !string.IsNullOrWhiteSpace(s))
                                                .LastOrDefault();
                    TimeSpan video_length = VideoParser.GetBrightcoveVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Bcove Link",
                                                video_id,
                                                link.InnerText + video_found,
                                                link.Attributes["href"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(link)));
                }
            }
        }
        private void ProcessIframes()
        {
            var iframe_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//iframe");
            if (iframe_list == null)
            {
                return;
            }
            foreach(var iframe in iframe_list)
            {
                string title = "";
                
                if(iframe.Attributes["title"] == null)
                {
                    title = "No title attribute found";
                }
                else
                {
                    title = iframe.Attributes["title"].Value;
                }

                if(iframe.Attributes["src"] == null)
                {
                    continue;
                }

                if (iframe.Attributes["src"].Value.Contains("youtube"))
                {
                    var uri = new Uri((iframe.Attributes["src"].Value));
                    var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
                    var video_id = string.Empty;
                    if (query.AllKeys.Contains("v"))
                    {
                        video_id = query["v"];
                    }
                    else
                    {
                        video_id = uri.Segments.LastOrDefault();
                    }
                    
                    TimeSpan video_length;
                    try
                    {
                        video_id = video_id.CleanSplit("?").FirstOrDefault();
                        video_length = VideoParser.GetYoutubeVideoLength(video_id);
                    }
                    catch
                    {
                        Console.WriteLine("Video not found");
                        video_length = new TimeSpan(0);
                    }
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "YouTube Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                true));
                }
                else if (iframe.Attributes["src"].Value.Contains("brightcove"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("=").LastOrDefault().CleanSplit("&")[0];
                    TimeSpan video_length = VideoParser.GetBrightcoveVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Brightcove Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("H5P"))
                {
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "H5P",
                                                "",
                                                title,
                                                iframe.Attributes["src"].Value,
                                                new TimeSpan(0),
                                                false));
                }else if (iframe.Attributes["src"].Value.Contains("byu.mediasite"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("/").LastOrDefault();
                    if (String.IsNullOrEmpty(video_id))
                    {
                        video_id = iframe.Attributes["src"].Value.CleanSplit("/").Reverse().Skip(1).FirstOrDefault();
                    }
                    TimeSpan video_length = VideoParser.GetBYUMediaSiteVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "BYU Mediasite Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("panopto"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("=").LastOrDefault().CleanSplit("&")[1];
                    TimeSpan video_length = VideoParser.GetPanoptoVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Panopto Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("alexanderstreet"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("token/").LastOrDefault();
                    TimeSpan video_length = VideoParser.GetAlexanderStreetVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "AlexanderStreet Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("kanopy"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("embed/").LastOrDefault();
                    TimeSpan video_length = VideoParser.GetKanopyVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Kanopy Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("ambrosevideo"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("?").LastOrDefault().CleanSplit("&")[0];
                    TimeSpan video_length = VideoParser.GetKanopyVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Ambrose Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("facebook"))
                {
                    string video_id = new Regex("\\d{17}").Match(iframe.Attributes["src"].Value).Value;
                    TimeSpan video_length = VideoParser.GetKanopyVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Facebook Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("dailymotion"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("/").LastOrDefault();
                    TimeSpan video_length = VideoParser.GetKanopyVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "DailyMotion Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else if (iframe.Attributes["src"].Value.Contains("vimeo"))
                {
                    string video_id = iframe.Attributes["src"].Value.CleanSplit("/").LastOrDefault().CleanSplit("?")[0];
                    TimeSpan video_length = VideoParser.GetKanopyVideoLength(video_id, Chrome, Wait);
                    string video_found;
                    if (video_length == new TimeSpan(0))
                    {
                        video_found = "\nVideo not found";
                    }
                    else
                    {
                        video_found = "";
                    }
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                "Vimeo Video",
                                                video_id,
                                                title + video_found,
                                                iframe.Attributes["src"].Value,
                                                video_length,
                                                VideoParser.CheckTranscript(iframe)));
                }
                else
                {
                    Data.Add(new PageMediaData(PageDocument.Location,
                                                 "Iframe",
                                                 "",
                                                 title,
                                                 iframe.Attributes["src"].Value,
                                                 new TimeSpan(0),
                                                 false));
                }
            }
        }
        private void ProcessVideoTags()
        {
            var video_tag_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//video");
            if (video_tag_list == null)
            {
                return;
            }

            foreach(var video in video_tag_list)
            {
                string video_id = video.Attributes["src"].Value.CleanSplit("=")[1].CleanSplit("&")[0];
                Data.Add(new PageMediaData(PageDocument.Location,
                                            "Inline Media Video",
                                            video_id,
                                            "Inline Media:\nUnable to find title or video length for this type of video",
                                            video.Attributes["src"].Value,
                                            new TimeSpan(0),
                                            VideoParser.CheckTranscript(video)));
            }
        }

        private void ProcessBrightcoveVideoHTML()
        {
            var brightcove_list = PageDocument.Doc
                .DocumentNode
                ?.SelectNodes(@"//div[@id]")
                ?.Where(e => new Regex("\\d{13}").IsMatch(e.Id));
            if (brightcove_list == null)
            {
                return;
            }
            foreach(var video in brightcove_list)
            {
                string video_id = new Regex("\\d{13}").Match(video.Id).Value;
                TimeSpan video_length = VideoParser.GetBrightcoveVideoLength(video_id, Chrome, Wait);
                Data.Add(new PageMediaData(PageDocument.Location,
                                            "Brightcove Video",
                                            video_id,
                                            "",
                                            $"https://studio.brightcove.com/products/videocloud/media/videos/search/{video_id}",
                                            video_length,
                                            VideoParser.CheckTranscript(video)));
            }
        }
    }
    public class LinkParser : RParserBase
    {
        //class to do a link report
        public LinkParser(string path)
        {
            Directory = path;
        }
        private string Directory = string.Empty;
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            if (page_info[page_info.Keys.ElementAt(0)] == null)
            {
                return;
            }
            PageDocument = new DataToParse(page_info.Keys.ElementAt(0), page_info[page_info.Keys.ElementAt(0)]);

            ProcessLinks();
            ProcesImages();
        }
        private bool TestUrl(string url)
        {

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "HEAD";
            request.Proxy = null;
            request.UseDefaultCredentials = true;
            try
            {
                using(HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    return response.StatusCode == HttpStatusCode.OK;
                }
            }
            catch(WebException)
            {
                return false;
            }
        }
        private bool TestPath(string path)
        {
            if(Directory == "None")
            {
                //Then we don't have a directory to compare against
                return true;
            }
            //Need to remove any HTML page location part of the file path as that will cause the test to fail.
            path = path.CleanSplit("#").FirstOrDefault();
            try
            {
                if (new Regex("^\\.\\.").IsMatch(path))
                {
                    path = Path.GetFullPath(Path.Combine(Directory, path));
                }
                else
                {
                    path = Path.GetFullPath(Path.Combine(Directory, path));
                }
            }
            catch
            {
                return false;
            }
            

            return File.Exists(path);
        }
        private void ProcessLinks()
        {
            var link_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//a");
            if (link_list == null)
            {
                return;
            }
            Parallel.ForEach(link_list, link =>
            {
                if(link.Attributes["href"] != null)
                {
                    if(new Regex("^#").IsMatch(link.Attributes["href"].Value))
                    {
                        //Do nothing
                    }
                    else if(new Regex("^mailto:").IsMatch(link.Attributes["href"].Value))
                    {
                        //Do nothing
                    }
                    else if(new Regex("^javascript:").IsMatch(link.Attributes["href"].Value))
                    {
                        Data.Add(new PageData(PageDocument.Location, 
                                                link.Attributes["href"].Value, 
                                                "", 
                                                "JavaScript links are often not accessible \\ broken."));
                    }
                    else if(new Regex("http|^www\\.|.*?\\.com$|.*?\\.org$").IsMatch(link.Attributes["href"].Value))
                    {
                        if (!TestUrl(link.Attributes["href"].Value))
                        {
                            Data.Add(new PageData(PageDocument.Location,
                                                    link.Attributes["href"].Value,
                                                    "",
                                                    "Broken link, needs to be checked"));
                        }
                    }
                    else
                    {
                        if (!TestPath(link.Attributes["href"].Value))
                        {
                            Data.Add(new PageData(PageDocument.Location,
                                                    link.Attributes["href"].Value,
                                                    "",
                                                    "File does not exist"));
                        }
                    }
                }
            });
        }
        private void ProcesImages()
        {
            var image_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//img");
            if (image_list == null)
            {
                return;
            }
            Parallel.ForEach(image_list, image => 
            {
                if(image.Attributes["src"] != null)
                {
                    if(new Regex("http|^www\\.|.*?\\.com$|.*?\\.org$").IsMatch(image.Attributes["src"].Value))
                    {

                    }
                    else
                    {
                        if (!TestPath(image.Attributes["src"].Value))
                        {
                            Data.Add(new PageData(PageDocument.Location,
                                                    image.Attributes["src"].Value,
                                                    "",
                                                    "Image does not exist"));
                        }
                    }
                }
            });
        }
    }
    public class CreateExcelReport
    {
        public CreateExcelReport(string destination_path)
        {
            this.Destination = destination_path;
            this.Excel = new ExcelPackage(new FileInfo(PathToExcelTemplate));
            this.Cells = Excel.Workbook.Worksheets[1].Cells;
            this.RowNumber = 9;
        }
        private string Destination;
        private string PathToExcelTemplate = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\AccessibilityTools\ReportGenerators-master\CAR - Accessibility Review Template.xlsx";
        private ExcelPackage Excel;
        private ExcelRange Cells;
        private int RowNumber;
        
        public void CreateReport(List<PageData> A11yData, List<PageData> MediaData, List<PageData> LinkData)
        {
            if (null != A11yData)
            {
                AddA11yData(A11yData);
            }
            if(null != MediaData)
            {
                AddMediaData(MediaData);
            }
            if (null != LinkData)
            {
                AddLinkData(LinkData);
            }
            var test_path = new DirectoryInfo(Path.GetDirectoryName(Destination));
            if (!(test_path.Exists))
            {
                test_path.Create();
            }
            var i = 1;
            while(new FileInfo(Destination).Exists)
            {
                var new_destination = Destination.Replace(".xlsx", $"_V{i}.xlsx");
                if(!(new FileInfo(new_destination).Exists))
                { 
                    Destination = new_destination;
                }
                i++;
            }
            Excel.SaveAs(new FileInfo(Destination));
            Excel.Dispose();
        }
        private void AddA11yData(List<PageData> data_list)
        {
            RowNumber = 9;
            Cells = Excel.Workbook.Worksheets[1].Cells;
            foreach(var data in data_list)
            {
                Cells[RowNumber, 2].Value = "Not Started";
                Cells[RowNumber, 3].Value = data.Location.CleanSplit("/").LastOrDefault().CleanSplit("\\").LastOrDefault();
                Cells[RowNumber, 3].Hyperlink =  new System.Uri(Regex.Replace(data.Location, "api/v\\d/", ""));
                switch ((data as PageA11yData).Issue.ToLower())
                {
                    case "adjust link text":
                        A11yAddToCell("Link", "Non-Descriptive Link", data.Text);
                        break;
                    case "javaScript links are not accessible":
                        A11yAddToCell("Link", "JavaScript Link", data.Text);
                        break;
                    case "broken link":
                        A11yAddToCell("Link", "Broken Link", data.Text);
                        break;
                    case "empty link tag":
                        A11yAddToCell("Link", "Broken Link", data.Text);
                        break;
                    case "needs a title":
                        A11yAddToCell("Semantics", "Missing title/label", $"{data.Element} needs a title attribute\nID: {data.Id}");
                        break;
                    case "no alt attribute":
                        A11yAddToCell("Image", "No Alt Attribute", data.Text);
                        break;
                    case "alt text may need adjustment":
                        A11yAddToCell("Image", "Non-Descriptive alt tags", data.Text);
                        break;
                    case "check if header is meant to be invisible and is not a duplicate":
                        A11yAddToCell("Semantics", "Improper Headings", $"Invisible header:\n{data.Text}");
                        break;
                    case "no transcript found":
                        A11yAddToCell("Media", "Transcript Needed", data.Text);
                        break;
                    case "revise table":
                        A11yAddToCell("Table", "", data.Text);
                        break;
                    case "<i>/<b> tags should be <em>/<strong> tags":
                        A11yAddToCell("Semantics", "Bad use of <i> and/or <b>", (data as PageA11yData).Issue);
                        break;
                    case "flash is inaccessible":
                        A11yAddToCell("Misc", "", $"{data.Text}\n{(data as PageA11yData).Issue}");
                        break;
                    case "does not meet aa color contrast":
                        A11yAddToCell("Color", "Doesn't meet contrast ratio", $"{(data as PageA11yData).Issue}\n{data.Text}");
                        break;
                    default:
                        A11yAddToCell("", "", (data.Element + "\n" + data.Text + "\n" + (data as PageA11yData).Issue));
                        break;
                }
                RowNumber++;
            }
        }
        private void A11yAddToCell(string issue_type, string descriptive_error, string notes, int severity = 1, int occurence = 1, int detection = 1)
        {
            Cells[RowNumber, 4].Value = issue_type;
            Cells[RowNumber, 5].Value = descriptive_error;
            Cells[RowNumber, 6].Value = notes;
            Cells[RowNumber, 7].Value = severity;
            Cells[RowNumber, 8].Value = occurence;
            Cells[RowNumber, 9].Value = detection;
        }
        private void AddMediaData(List<PageData> data_list)
        {
            //I may need to do a replace on the media item url, 
            //#Sometimes there are extra slashes at the beginning
            //$url = $url - replace "^//", "https://"
            RowNumber = 4;
            Cells = Excel.Workbook.Worksheets[2].Cells;
            Excel.Workbook.Worksheets[2].Column(4).Style.Numberformat.Format = "#############";
            Excel.Workbook.Worksheets[2].Column(6).Style.Numberformat.Format = "hh:mm:ss";
            Excel.Workbook.Worksheets[2].Column(11).Style.Numberformat.Format = "hh:mm:ss";
            Excel.Workbook.Worksheets[2].Column(12).Style.Numberformat.Format = "hh:mm:ss";
            foreach (var data in data_list)
            {
                Cells[RowNumber, 2].Value = data.Element;
                Cells[RowNumber, 3].Value = data.Location.CleanSplit("/").LastOrDefault().CleanSplit("\\").LastOrDefault();
                Cells[RowNumber, 3].Hyperlink = new System.Uri(data.Location);
                if((from cell in Cells["D:D"] where cell.Value?.ToString() == data.Id select true).Count(c => c == true) > 0)
                {
                    Cells[RowNumber, 4].Value = "Duplicate Video:\n" + data.Id;
                }
                else
                {
                    Cells[RowNumber, 4].Value = data.Id;
                }
                Cells[RowNumber, 5].Value = Regex.Replace((data as PageMediaData).MediaUrl, "^//", "https://");
                Cells[RowNumber, 5].Hyperlink = new System.Uri(Regex.Replace((data as PageMediaData).MediaUrl, "^//", "https://"));
                Cells[RowNumber, 6].Value = (data as PageMediaData).VideoLength;
                Cells[RowNumber, 7].Value = data.Text;
                Cells[RowNumber, 8].Value = (data as PageMediaData).Transcript ? "Yes" : "No";
                RowNumber++;
            }
        }
        
        private void AddLinkData(List<PageData> data_list)
        {
            Cells = Excel.Workbook.Worksheets[3].Cells;
            RowNumber = 4;
            foreach (var data in data_list)
            {
                Cells[RowNumber, 2].Value = data.Location.CleanSplit("/").LastOrDefault().CleanSplit("\\").LastOrDefault();
                Cells[RowNumber, 2].Hyperlink = new System.Uri(data.Location);
                Cells[RowNumber, 3].Value = data.Element;
                if (data.Element.Contains("http"))
                {
                    Cells[RowNumber, 3].Hyperlink = new System.Uri(data.Element);
                }
                Cells[RowNumber, 4].Value = data.Text;
                RowNumber++;
            }
        }
    }
    public class GenerateReport
    {
        //This is where the program will start and take user input / run the reports, may or may not be needed based on how I can get the SpecFlow test to work.
        public static void Main()
        {
            CourseInfo course = new CourseInfo(@"I:\Canvas\FOODS-043\FOODS-043-S003\HTML");
            LinkParser parser = new LinkParser(course.CourseIdOrPath);
            foreach(var page in course.PageHtmlList)
            {
                parser.ProcessContent(page);
            }
            CreateExcelReport GenReport = new CreateExcelReport(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\AccessibilityTools\\ReportGenerators-master\\Reports\\ARC_{course.CourseCode}_{CanvasApi.CurrentDomain}.xlsx");
            GenReport.CreateReport(null, null, parser.Data);
        }
    }
}
