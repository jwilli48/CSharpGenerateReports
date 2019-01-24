namespace ReportGenerators
{
    using System;
    using System.Collections.Generic;
    using HtmlAgilityPack;
    using System.Threading.Tasks;
    using My.CanvasApi;

    public class DataToParse
    {   //Object stored in the ReportParser objects that turns the html string from the CourseInfo object into a live HTML dom to be used by the parsers.
        public DataToParse(string location, string page_body)
        {
            Location = location;
            Doc = new HtmlDocument();
            Doc.LoadHtml(page_body);
        }
        public DataToParse(string location, HtmlDocument doc)
        {
            Location = location;
            Doc = doc;
        }
        public string Location;
        public HtmlDocument Doc;
    }

    public abstract class RParserBase
    {
        //Base class for each of the reports
        public RParserBase() { }
        public List<PageData> Data { get; set; } = new List<PageData>();
        public abstract void ProcessContent(Dictionary<string, string> page_info);

    }
   
    public class GenerateReport
    {
        //This is where the program will start and take user input / run the reports, may or may not be needed based on how I can get the SpecFlow test to work.
        //It is currently just a testing function
        public static void Main()
        {
            var s = new System.Diagnostics.Stopwatch();
            s.Start();

            CanvasApi.ChangeDomain("Directory");

            CourseInfo course = new CourseInfo(@"I:\Canvas\USA-043\USA-043-002\HTML");
            bool directory = true;
            LinkParser ParseForLinks = new LinkParser(@"I:\Canvas\USA-043\USA-043-002\HTML"); //Need to declare this early as it is only set if it is a directory

            A11yParser ParseForA11y = new A11yParser();
            MediaParser ParseForMedia = new MediaParser();

            Parallel.ForEach(course.PageHtmlList, page =>
            {

                ParseForA11y.ProcessContent(page);
                ParseForMedia.ProcessContent(page);
                if (directory)
                {
                    ParseForLinks.ProcessContent(page);
                }
            });

            CreateExcelReport GenReport = new CreateExcelReport(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\AccessibilityTools\\ReportGenerators-master\\Reports\\ARC_{course.CourseCode}_{CanvasApi.CurrentDomain}.xlsx");
            GenReport.CreateReport(ParseForA11y.Data, ParseForMedia.Data, ParseForLinks?.Data);
            s.Stop();
            ParseForMedia.Chrome.Quit();
            Console.Write(ParseForLinks.Time.TotalSeconds);
        }
    }
}
