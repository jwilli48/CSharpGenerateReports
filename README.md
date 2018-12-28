# CSharpGenerateReports
Mostly a learning project to convert my POSH code to C#

This should actually now work much better then the PowerShell version now. The only thing it needs for anyone to get it working is to add code to the main at the bottom of the RGeneratorBase.cs file to get any user input / however they want to get the course ID's or paths and to then run the parsers on it. It would look something like this:
    
    var id = Console.ReadLine();
    CourseInfo course = new CourseInfo(id);
    
    A11yParser ParseForA11y = new A11yParser();
    MediaParser ParseForMedia = new MediaParser();
    LinkParser ParseForLinks = new LinkParser(course.CourseIdOrPath);
    foreach(var page in course.PageHtmlList){
        ParseForA11y.ProceessContent(page);
        ParseForMedia.ProcessContent(page);
        ParseForLinks.ProcessContent(page);
    }
    
    CreateExcelReport GenReport = new CreateExcelReport(@"C:\Users\UserName\Temp\Test.xlsx");
    GenReport.CreateReport(ParseForA11y.Data, ParseForMedia.Data, ParseForLinks.Dara);
