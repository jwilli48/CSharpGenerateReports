using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ReportGenerators
{
    public class CourseInfo
    {
        //Class that will contain all of the courses info, including the URL and HTML body for each item
        //This CourseInfo class will also use the CanvasAPI static class in order to fill it with the course information upon construction (this is the longest time consuming part of the program)
        //Can't really make it multithreaded as the CanvasAPI has an access limit that would be used up to quickly if we run multiple requests at the same time.
        //Should maybe make the fill object a mehtod instead of just inside the constructor.
        public CourseInfo(string course_path)
        {
            //Constructor for if a directory path is input
            this.CourseIdOrPath = course_path;
            string[] array = course_path.Split('\\');
            this.CourseName = array.Take(array.Length - 1).LastOrDefault();
            this.CourseCode = array.Take(array.Length - 1).LastOrDefault();

            PageHtmlList = new List<Dictionary<string, string>>();
            foreach(var file in Directory.GetFiles(course_path, "*.html", SearchOption.TopDirectoryOnly))
            {
                Console.WriteLine(file.CleanSplit("\\").LastOrDefault());
                string location = string.Empty;
                switch (Path.GetPathRoot(file))
                {
                    case "I:\\":
                        location = $"https://iscontent.byu.edu/{file.Replace("I:\\", "")}";
                        break;
                    case "Q:\\":
                        location = $"https://isdev.byu.edu/courses/{file.Replace("Q:\\", "")}";
                        break;
                    default:
                        location = $"file:///{file}";
                        break;
                }
                var temp_dict = new Dictionary<string, string>
                {
                    [location] = File.ReadAllText(file)
                };
                PageHtmlList.Add(temp_dict);
            }
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
            foreach (CanvasModule module in CanvasApi.GetCanvasModules(course_id))
            {
                Console.WriteLine(module.name);
                //Loop through all the items for each module
                foreach (CanvasModuleItem item in CanvasApi.GetCanvasModuleItems(course_id, module.id))
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
                                        LocationAndBody[item.url] += "\n" + question.question_text;
                                        //Loop through all answers in the quiz
                                        foreach (CanvasQuizQuestionAnswers answer in question.answers)
                                        {
                                            LocationAndBody[item.url] += "\n" + answer.html;
                                            LocationAndBody[item.url] += "\n" + answer.comments_html;
                                        }
                                    }
                                }
                                catch (Exception e)
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
                                LocationAndBody["Empty"] = null;
                                break;
                        }
                        //Add the location and HTML body to the List
                        PageHtmlList.Add(LocationAndBody);
                    }
                    catch (Exception e)
                    {
                        //Check if it was unauthorized
                        if (e.Message.Contains("Unauthorized"))
                        {
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
        public List<Dictionary<string, string>> PageHtmlList { get; set; }
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
            foreach (var p in props)
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
        public PageMediaData(string location, string element, string id, string text, string media_url, TimeSpan video_length, bool transcript) : base(location, element, id, text)
        {
            this.MediaUrl = media_url;
            this.VideoLength = video_length;
            this.Transcript = transcript;
        }
        public string MediaUrl { get; }
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
}
