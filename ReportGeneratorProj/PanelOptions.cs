using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace My
{
    public class PanelOptions
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string ChromeDriverPath { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string FirefoxDriverPath { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string QDriveContentUrl { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string IDriveContentUrl { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        Dictionary<string, string> BrightCoveCred { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string JsonDataDir { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        Dictionary<string, string> BYUOnlineCreds { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        Dictionary<string, string> BYUISTestCreds { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        Dictionary<string, string> BYUMasterCoursesCreds { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        Dictionary<string, string> ByuCred { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string ReportPath { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string PowershellScriptDir { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string ExcelTemplatePath { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string GoogleApi { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        List<string> FilesToIgnore { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string HighScorePath { get; set; }
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        string CourseBackupDir { get; set; }
    }
}
