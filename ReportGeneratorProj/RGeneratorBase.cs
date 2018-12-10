using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using RestSharp;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
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
    }
    public class DataToParse
    {
        public DataToParse(string location, string page_body)
        {
            this.Location = location;
            Doc = new HtmlDocument();
            Doc.Load(page_body);
        }
        public string Location;
        public HtmlDocument Doc;
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
        A11yParser() { }
        public override void ProcessContent(Dictionary<string, string> page_info)
        {
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
                if (link.InnerText.Contains("<img"))
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
            
        }
        private void ProcessTables()
        {

        }
        private void ProcessIframes()
        {

        }
        private void ProcessBrightcoveVideoHTML()
        {

        }
        private void ProcessHeaders()
        {

        }
        private void ProcessSemantics()
        {

        }
        private void ProcessVideoTags()
        {

        }
        private void ProcessFlash()
        {

        }
        private void ProcessColor()
        {

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

    public class GenerateReport
    {
        //This is where the program will start and take user input / run the reports, may or may not be needed based on how I can get the SpecFlow test to work.
        public static void Main()
        {
            CourseInfo course = new CourseInfo(1026);
            Console.ReadLine();
        }
    }
}
