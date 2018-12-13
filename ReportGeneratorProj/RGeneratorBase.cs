using System;
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

namespace ReportGenerators
{
    //Name space for all classes needed for the ReportGenerators
    public class CanvasCourse
    {
        //https://canvas.instructure.com/doc/api/courses.html
        public int id { get; set; }
        public string name { get; set; }
        public string course_code { get; set; }
    }
    public class CanvasModule
    {
        //https://canvas.instructure.com/doc/api/modules.html
        public int id { get; set; }
        public string name { get; set; }
        public int items_count { get; set; }
        public string items_url { get; set; }
    }
    public class CanvasModuleItem
    {
        //https://canvas.instructure.com/doc/api/modules.html
        public int id { get; set; }
        public string title { get; set; }
        public string type { get; set; }
        public int content_id { get; set; }
        public string html_url { get; set; }
        public string url { get; set; }
        public string page_url { get; set; }
    }
    public class CanvasPage
    {
        //https://canvas.instructure.com/doc/api/pages.html
        public string url { get; set; }
        public string title { get; set; }
        public string body { get; set; }
    }
    public class CanvasDiscussionTopic
    {
        //https://canvas.instructure.com/doc/api/discussion_topics.html
        public int id { get; set; }
        public string title { get; set; }
        public string message { get; set; }
        public string html_url { get; set; }
    }
    public class CanvasAssignment
    {
        //https://canvas.instructure.com/doc/api/assignments.html
        public int id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string html_url { get; set; }

    }
    public class CanvasQuiz
    {
        //https://canvas.instructure.com/doc/api/quizzes.html
        public int id { get; set; }
        public string title { get; set; }
        public string description { get; set; }

    }
    public class CanvasQuizQuestionAnswers
    {
        //https://canvas.instructure.com/doc/api/quiz_questions.html
        public int id { get; set; }
        public string answer_text { get; set; }
        public string answer_comments { get; set; }
        public string html { get; set; }
        public string comments_html { get; set; }
    }
    public class CanvasQuizQuesiton
    {
        //https://canvas.instructure.com/doc/api/quiz_questions.html
        public int id { get; set; }
        public string question_name { get; set; }
        public string question_type { get; set; }
        public string question_text { get; set; }
        public List<CanvasQuizQuestionAnswers> answers{ get; set;}
    }
    public class CanvasApi
    {
        //Class to control interaction with the Canvas API
        //Token is needed to authenticate, will need to adjust this to be stored in a seperate file instead of in this code.
        private const string token = "";
        //THe base domain url for the API
        //BYU has 3 main domain names
        private const string domain = "byu.instructure.com";
        public static CanvasCourse GetCanvasCourse(int course_id)
        {
            //Will send request for basic course information
            string url = $"https://{domain}/api/v1/courses/{course_id}?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            //Will return single course object with parameters we want
            var response = restClient.Execute<CanvasCourse>(request);
            return response.Data;
        }
        public static List<CanvasModule> GetCanvasModules (int course_id)
        {
            //Request for all modules within a course
            string url = $"https://{domain}/api/v1/courses/{course_id}/modules?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            //Returns a List of CanvasModule objects
            var response = restClient.Execute<List<CanvasModule>>(request);
            return response.Data;
        }
        public static List<CanvasModuleItem> GetCanvasModuleItems (int course_id, int module_id)
        {
            //Request for all items within a module
            string url = $"https://{domain}/api/v1/courses/{course_id}/modules/{module_id}/items?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            //Returns a List of CanvasModuleItems
            var response = restClient.Execute<List<CanvasModuleItem>>(request);
            return response.Data;
        }
        public static CanvasPage GetCanvasPage (int course_id, string page_url)
        {
            //Request for a single canvas page
            string url = $"https://{domain}/api/v1/courses/{course_id}/pages/{page_url}?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<CanvasPage>(request);
            return response.Data;
        }
        public static CanvasDiscussionTopic GetCanvasDiscussionTopics(int course_id, int topic_id)
        {
            //Request for a signle discussion topic
            string url = $"https://{domain}/api/v1/courses/{course_id}/discussion_topics/{topic_id}?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<CanvasDiscussionTopic>(request);
            return response.Data;
        }
        public static CanvasAssignment GetCanvasAssignments(int course_id, int content_id)
        {
            //Request for a single canvas assignment page
            string url = $"https://{domain}/api/v1/courses/{course_id}/assignments/{content_id}?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<CanvasAssignment>(request);
            return response.Data;
        }
        public static CanvasQuiz GetCanvasQuizzes(int course_id, int content_id)
        {
            //Request for a single canvas quiz
            string url = $"https://{domain}/api/v1/courses/{course_id}/quizzes/{content_id}?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<CanvasQuiz>(request);
            return response.Data;
        }
        public static List<CanvasQuizQuesiton> GetCanvasQuizQuesitons(int course_id, int content_id)
        {
            //Request for the list of quiz questions (and answers) of a canvas quiz
            string url = $"https://{domain}/api/v1/courses/{course_id}/quizzes/{content_id}/questions?per_page=10000&access_token={token}";
            var restClient = new RestClient(url);
            var request = new RestRequest(Method.GET);
            var response = restClient.Execute<List<CanvasQuizQuesiton>>(request);
            return response.Data;
        }

    }
    public class CourseInfo
    {
        //Class that will contain all of the courses info, including the URL and HTML body for each item
        public CourseInfo(string course_path)
        {
            //Constructor for if a directory path is input
            this.CourseIdOrPath = course_path;
            string[] array = course_path.Split('\\');
            this.CourseName = array.Take(array.Length - 1).LastOrDefault();
            this.CourseCode = array.Take(array.Length - 1).LastOrDefault();

            PageHtmlList = new List<Dictionary<string, string>>();
        }
        public CourseInfo(int course_id)
        {
            //Constructor for a canvas course ID number
            this.CourseIdOrPath = course_id;
            //Get the course information with API
            CanvasCourse course_info = CanvasApi.GetCanvasCourse(course_id);
            CourseName = course_info.name;
            CourseCode = course_info.course_code;
            //Need to make sure the HtmlList is initialized so we can store all of the info
            PageHtmlList = new List<Dictionary<string, string>>();
            //Begin to loop through all modules of the course
            foreach(CanvasModule module in CanvasApi.GetCanvasModules(course_id))
            {
                Console.WriteLine(module.name);
                //Loop through all the items for each module
                foreach(CanvasModuleItem item in CanvasApi.GetCanvasModuleItems(course_id, module.id))
                {
                    //The object to connect the item location and its HTML body
                    Dictionary<string, string> LocationAndBody = new Dictionary<string, string>();
                    Console.WriteLine(item.title);
                    try //This try block is just in case we are not authroized to access any of these pages
                    {
                        switch (item.type)
                        { //Need to see what type of item it is to determine request needed
                            case "Page":
                                CanvasPage page = CanvasApi.GetCanvasPage(course_id, item.page_url);
                                LocationAndBody[item.url] = page.body;
                                break;
                            case "Discussion":
                                CanvasDiscussionTopic discussion = CanvasApi.GetCanvasDiscussionTopics(course_id, item.content_id);
                                LocationAndBody[item.url] = discussion.message;
                                break;
                            case "Assignment":
                                CanvasAssignment assignment = CanvasApi.GetCanvasAssignments(course_id, item.content_id);
                                LocationAndBody[item.url] = assignment.description;
                                break;
                            case "Quiz":
                                CanvasQuiz quiz = CanvasApi.GetCanvasQuizzes(course_id, item.content_id);
                                LocationAndBody[item.url] = quiz.description;
                                try
                                { //Quizes require more as we need to gather question and answer info
                                    //Again may be able to see basic quiz but not authorized for quiz questions, this the try block.
                                    //Loop through all questions for specific quiz
                                    foreach (CanvasQuizQuesiton question in CanvasApi.GetCanvasQuizQuesitons(course_id, item.content_id))
                                    {
                                        LocationAndBody[item.url] += question.question_text;
                                        //Loop through all answers in the quiz
                                        foreach (CanvasQuizQuestionAnswers answer in question.answers)
                                        {
                                            LocationAndBody[item.url] += answer.html;
                                            LocationAndBody[item.url] += answer.comments_html;
                                        }
                                    }
                                }catch(Exception e)
                                {
                                    //Check if the exception was an unauthorized request
                                    if (e.Message.Contains("Unauthorized"))
                                    {
                                        Console.WriteLine("ERROR: (401) Unauthorized, can not search quiz questions. Skipping...");
                                    }
                                    else
                                    {
                                        Console.WriteLine("{0}", e);
                                    }
                                }
                                break;
                            default:
                                Console.WriteLine($"Not Supported:\n{item.type}");
                                break;
                        }
                        //Add the location and HTML body to the List
                        this.PageHtmlList.Add(LocationAndBody);
                    }
                    catch(Exception e)
                    {
                        //Check if it was unauthorized
                        if(e.Message.Contains("Unauthorized")){
                            Console.WriteLine($"ERROR: (401) Unauthorized, can not search:\n{item.title}\n{item.type}");
                        }
                        else
                        {
                            Console.WriteLine("{0}", e);
                        }
                    }
                }
            }
        }
        public dynamic CourseIdOrPath { get; }
        public string CourseName { get; }
        public string CourseCode { get; }
        public List<Dictionary<string,string>> PageHtmlList { get; set; }
    }
    public class PageData
    {
        //Base clas for holding issues / data from a single page
        public PageData(string input_location, string input_element, string input_id, string input_text)
        {
            this.Location = input_location;
            this.Element = input_element;
            this.Id = input_id;
            this.Text = input_text;
        }
        public string Location { get; }
        public string Element { get; }
        public string Id { get; }
        public string Text { get; }
        public override string ToString()
        {
            var props = typeof(PageData).GetProperties();
            var sb = new StringBuilder();
            foreach(var p in props)
            {
                sb.AppendLine(p.Name + ": " + p.GetValue(this, null));
            }
            return sb.ToString();
        }
    }
    public class PageA11yData : PageData
    {
        //Exxtension of class for accessibility params desired
        public PageA11yData(string location, string element, string id, string text, string issue, int severity) : base(location, element, id, text)
        {
            this.Issue = issue;
            this.Severity = severity;
        }
        public string Issue { get; }
        public int Severity { get; }
        public override string ToString()
        {
            var props = typeof(PageA11yData).GetProperties();
            var sb = new StringBuilder();
            foreach (var p in props)
            {
                sb.AppendLine(p.Name + ": " + p.GetValue(this, null));
            }
            return sb.ToString();
        }
    }
    public class PageMediaData : PageData
    {
        //Extension of class for Media data from a page
        public PageMediaData(string location, string element, string id, string text, Uri media_url, TimeSpan video_length, bool transcript) : base(location, element, id, text)
        {
            this.MediaUrl = media_url;
            this.VideoLength = video_length;
            this.Transcript = transcript;
        }
        public Uri MediaUrl { get; }
        public TimeSpan VideoLength { get; }
        public bool Transcript { get; }
        public override string ToString()
        {
            var props = typeof(PageMediaData).GetProperties();
            var sb = new StringBuilder();
            foreach (var p in props)
            {
                sb.AppendLine(p.Name + ": " + p.GetValue(this, null));
            }
            return sb.ToString();
        }
    }
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
    public class VideoParser
    {
        public static bool CheckTranscript(HtmlNode element)
        {
            if(element.NextSibling.OuterHtml.Contains("transcript") || element.NextSibling.NextSibling.OuterHtml.Contains("transcript"))
            {
                return true;
            }
            return false;
        }
    }
    public abstract class RParserBase
    {
        //Base class for each of the reports
        //Due to there being multiple possible inputs to parse through, need to have a constructor for each type.
        public RParserBase() { }
        public List<PageData> Data { get; set; } = new List<PageData>();
        public DataToParse PageDocument;
        public abstract void ProcessContent(Dictionary<string, string> page_info);

    }
    public class A11yParser : RParserBase
    {
        //Class to do an accessibiltiy report
        public A11yParser() { }
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            if(page_info[page_info.Keys.ElementAt(0)] == null){
                return;
            }
            PageDocument = new DataToParse(page_info.Keys.ElementAt(0), page_info[page_info.Keys.ElementAt(0)]);

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
            var link_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//a");
            if(link_list == null)
            {
                return;
            }
            foreach(var link in link_list)
            {
                if(link.Attributes["onclick"] != null)
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.OuterHtml, "JavaScript links are not accessible", 1));
                }
                else if(link.Attributes["href"] == null)
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.OuterHtml, "Empty link tag", 1));
                }
                if (link.InnerHtml.Contains("<img"))
                {
                    continue;
                }
                if(link.InnerText == null)
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", "Invisible link with no text", "Adjust Link Text", 1));
                }else if(new Regex("^ ?here").IsMatch(link.InnerText))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                }else if(new Regex("^ ?[A-Za-z\\.]+ ?$").IsMatch(link.InnerText))
                {
                    if(link_list.Where(s => s.InnerText == link.InnerText).Count() > 1)
                    {
                        Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                    }
                }else if(new Regex("http|www\\.|Link|Click").IsMatch(link.InnerText))
                {
                    if(new Regex("Links to an external site").IsMatch(link.InnerText))
                    {
                        continue;
                    }
                    Data.Add(new PageA11yData(PageDocument.Location, "Link", "", link.InnerText, "Adjust Link Text", 1));
                }
            }
        }
        private void ProcessImages()
        {
            var image_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//img");
            if(image_list == null)
            {
                return;
            }
            foreach(var image in image_list)
            {
                var alt = image.Attributes["alt"]?.Value;
                if (alt == null)
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", image.OuterHtml, "No alt attribute", 1));
                }
                else if (new Regex("banner").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if (new Regex("Placeholder").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if (new Regex("\\.jpg").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("\\.png").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("http").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
                else if(new Regex("LaTeX:").IsMatch(alt))
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Image", "", alt, "Alt text may need adjustment", 1));
                }
            }
        }
        private void ProcessTables()
        {
            var table_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//table");
            if(table_list == null)
            {
                return;
            }
            var table_num = 1;
            foreach(var table in table_list)
            {
                var table_headers = table.SelectNodes(".//th");
                var table_data_cells = table.SelectNodes(".//td");
                var table_rows = table.SelectNodes(".//tr");
                var stretched_cells = table.SelectNodes(".//*[@colspan]");
                string issues = "";
                if(stretched_cells != null)
                {
                    issues += "Stretched table cell(s) should be a <caption> title for the table";
                }
                var num_rows = table_rows.Count();
                if(num_rows >= 3)
                {
                    if(table_headers == null)
                    {
                        issues += "\nTable has no headers";
                    }
                }
                var scope_headers = table_headers?.Count(c => c.Attributes["scope"] != null);
                if(scope_headers == null || scope_headers <= table_headers.Count())
                {
                    issues += "\nTable headers should have a scope attribute";
                }
                var scope_cells = table_data_cells?.Count(c => c.Attributes["scope"] != null);
                if(scope_cells != null && scope_cells > 0)
                {
                    issues += "\nNon-header table cells should not have scope attributes";
                }
                if(issues != null && issues != "")
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Table", "", $"Table number {table_num}:`n{issues}", "Revise table", 1));
                }
                table_num++;
            }
        }
        private void ProcessIframes()
        {
            var iframe_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//iframe");
            if(iframe_list == null)
            {
                return;
            }
            var iframe_number = 1;
            foreach(var iframe in iframe_list)
            {
                var src = iframe.Attributes["src"].Value;
                if (iframe.Attributes["title"] == null)
                {
                    if(new Regex("youtube").IsMatch(src))
                    {
                        var uri = new Uri(src);
                        var query = System.Web.HttpUtility.ParseQueryString(uri.Query);
                        var videoId = string.Empty;
                        if (query.AllKeys.Contains("v"))
                        {
                            videoId = query["v"];
                        }
                        else
                        {
                            videoId = uri.Segments.Last();
                        }
                        Data.Add(new PageA11yData(PageDocument.Location, "Youtube Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("brightcove").IsMatch(src))
                    {
                        var videoId = src.Split('=').Last().Split('&')[0];
                        if (!src.Contains("https:"))
                        {
                            src = $"https:{src}";
                        }
                        Data.Add(new PageA11yData(PageDocument.Location, "Brightcove Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("H5P").IsMatch(src))
                    {
                        Data.Add(new PageA11yData(PageDocument.Location, "H5P", "", "", "Needs a title", 1));
                    }else if(new Regex("byu\\.mediasite").IsMatch(src))
                    {
                        var videoId = src.Split('/').Last();
                        Data.Add(new PageA11yData(PageDocument.Location, "BYU Mediasite Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("panopto").IsMatch(src))
                    {
                        var videoId = src.Split('=').Last().Split('&')[1];
                        Data.Add(new PageA11yData(PageDocument.Location, "Panopto Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("alexanderstreet").IsMatch(src))
                    {
                        var videoId = src.Split(new string[] { "token/" }, StringSplitOptions.None).Last();
                        Data.Add(new PageA11yData(PageDocument.Location, "AlexanderStreen Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("kanopy").IsMatch(src))
                    {
                        var videoId = src.Split(new string[] { "embed/" }, StringSplitOptions.None).Last();
                        Data.Add(new PageA11yData(PageDocument.Location, "Kanopy Video", videoId, "", "Needs a title", 1));
                    }
                    else if(new Regex("ambrosevideo").IsMatch(src))
                    {
                        var videoId = src.Split('?').Last().Split('&')[0];
                        Data.Add(new PageA11yData(PageDocument.Location, "Ambrose Video", videoId, "", "NEeds a title", 1));
                    }else if(new Regex("facebook").IsMatch(src))
                    {
                        var videoId = new Regex("\\d{17}").Match(src).Value;
                        Data.Add(new PageA11yData(PageDocument.Location, "Facebook Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("dailymotion").IsMatch(src))
                    {
                        var videoId = src.Split('/').Last();
                        Data.Add(new PageA11yData(PageDocument.Location, "Facebook Video", videoId, "", "Needs a title", 1));
                    }else if(new Regex("vimeo").IsMatch(src))
                    {
                        var videoId = src.Split('/').Last().Split('?')[0];
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
                .SelectNodes("//div[matches(@id, '\\d{13}')");
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
                if (header.Attributes["class"].Value.Contains("screenreader-only"))
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
                var videoId = src.Split('=')[1].Split('&')[0];
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
                .SelectNodes("//object[contains(@id, \"flash\"");
            if(flash_list == null)
            {
                return;
            }
            else if(flash_list.Count() > 0)
            {
                Data.Add(new PageA11yData(PageDocument.Location, "Flash Element", "", $"{flash_list.Count()} embedded flash element(s) on this page", "Flash is inaccessible", 5));
            }

        }
        private void ProcessColor()
        {
            var colored_element_list = PageDocument.Doc
                .DocumentNode
                .SelectNodes("//*[contains(@style, \"color\"");
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
                {
                    background_color = "FFFFFF";
                }
                var foreground_color = style["color"];
                if(foreground_color == null)
                {
                    foreground_color = "000000";
                }
                if (!background_color.Contains("#"))
                {
                    int colorValue = System.Drawing.Color.FromName(background_color).ToArgb();
                    background_color = string.Format("{0:x6}", colorValue);
                }
                if (!foreground_color.Contains('#'))
                {
                    int colorValue = System.Drawing.Color.FromName(foreground_color).ToArgb();
                    foreground_color = string.Format("{0:x6}", colorValue);
                }
                foreground_color = foreground_color.Replace("#", "");
                background_color = background_color.Replace("#", "");
                var restClient = new RestClient($"https://webaim.org/resources/contrastchecker/?fcolor={foreground_color}&bcolor={background_color}&api");
                var request = new RestRequest(Method.GET);
                //Will return single course object with parameters we want
                var response = restClient.Execute<ColorContrast>(request).Data;
                if(response.AA != "pass")
                {
                    Data.Add(new PageA11yData(PageDocument.Location, "Color Contrast", "", $"Color: {foreground_color}\nBackgroundColor: {background_color}\n{response.ToString()}", "Does not meet AA color contrast", 1));
                }
            }
        }
    }
    public class ColorContrast
    {
        public double ratio { get; set;  }
        public string AA { get; set;  }
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
    public class MediaParser : RParserBase
    {
        //Class to do a media report
        public MediaParser()
        {
            Chrome = new ChromeDriver(@"E:\SeleniumTest");
            Wait = new WebDriverWait(Chrome, new TimeSpan(0, 0, 5));
        }
        //Gen a media report
        private ChromeDriver Chrome { get; set; }
        private WebDriverWait Wait { get; set; }
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            throw new NotImplementedException();
        }
    }
    public class LinkParser : RParserBase
    {
        //class to do a link report
        public LinkParser() { }
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
            throw new NotImplementedException();
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
        private string PathToExcelTemplate = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\CAR - Accessibility Review Template.xlsx";
        private ExcelPackage Excel;
        private ExcelRange Cells;
        private int RowNumber;
        private void A11yAddToCell(string issue_type, string descriptive_error, string notes, int severity = 1, int occurence = 1, int detection = 1)
        {
            Cells[RowNumber, 4].Value = issue_type;
            Cells[RowNumber, 5].Value = descriptive_error;
            Cells[RowNumber, 6].Value = notes;
            Cells[RowNumber, 7].Value = severity;
            Cells[RowNumber, 8].Value = occurence;
            Cells[RowNumber, 9].Value = detection;
        }
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
                Destination = Destination.Replace(".xlsx", $"_V{i}.xlsx");
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
                Cells[RowNumber, 2].Value = data.Location;
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
        private void AddMediaData(List<PageData> data_list)
        {

        }
        private void AddLinkData(List<PageData> data_list)
        {

        }
    }
    public class GenerateReport
    {
        //This is where the program will start and take user input / run the reports, may or may not be needed based on how I can get the SpecFlow test to work.
        public static void Main()
        {
            CourseInfo course = new CourseInfo(1026);
            A11yParser ParseForA11y = new A11yParser();
            foreach (var page in course.PageHtmlList)
            {
                ParseForA11y.ProcessContent(page);
            }
            var Destination = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + $"\\Reports\\ARC_{course.CourseCode}.xlsx";
            CreateExcelReport GenReport = new CreateExcelReport(Destination);
            GenReport.CreateReport(ParseForA11y.Data, null, null);
            
            Console.ReadLine();
        }
    }
}
