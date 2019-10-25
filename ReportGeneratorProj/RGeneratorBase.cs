﻿namespace ReportGenerators
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
            A11yParser a11YParser = new A11yParser();
            CourseInfo course = new CourseInfo(@"Q:\BrainHoney\Courses\ACC-200\ACC-200-M001\HTML");


        }
    }
}
