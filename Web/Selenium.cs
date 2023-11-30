using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.Extensions;
using RpaLib.ProcessAutomation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using RpaLib.Tracing;
using System.Diagnostics;

/// General Selenium class to ease manipulation of selenium driver
/// and initial configurations applied to all projects
namespace RpaLib.Web
{
    public static class Selenium
    {
        public static string Bin { get; private set; }
        public static string CurrentVersion { get; private set; }
        public static ChromeDriver Driver { get; private set; }
        // Connects to browser and disable download file extension protection
        public static ChromeDriver Connect(string chromedriverPath = null, string chromeBinPath = null, string defDwdDir = @"%USERPROFILE%\Downloads", int implicitWait = 500)
        {
            ChromeOptions options = new ChromeOptions();
            // To disable "This file can harm your computer..." message
            options.AddUserProfilePreference("safebrowsing.enabled", "false");
            options.AddUserProfilePreference("download.default_directory", Ut.GetFullPath(defDwdDir));
            if (chromeBinPath != null)
                options.BinaryLocation = chromeBinPath;
            //options.AddUserProfilePreference("download.prompt_for_download", "false");
            //options.AddUserProfilePreference("disable-popup-blocking", "true");
            try
            {
                // if called without specifying the driverpath, just try to run command in cmd
                if (chromedriverPath == null)
                {
                    Bin = "chromedriver.exe";
                    Driver = new ChromeDriver(options);
                }
                else
                {
                    Bin = Path.Combine(chromedriverPath, "chromedriver.exe");
                    var chromedriverService = ChromeDriverService.CreateDefaultService(chromedriverPath);
                    Driver = new ChromeDriver(chromedriverService, options);
                }

                //Log.Write(Rpa.RunPromptCommand("chromedriver", "--version"));
                Log.Write("ChromeDriver Version: " + GetDriverVersion());
            }
            catch (InvalidOperationException ex)
            {
                if (Ut.IsMatch(ex.Message, @"session not created: This version of ChromeDriver only supports Chrome version \d+"))
                {
                    Log.Write("Killing all \"chromedriver\" processes...");
                    foreach (Process proc in Process.GetProcessesByName("chromedriver"))
                        proc.Kill();

                    throw new ChromeDriverDeprecatedException(ex.Message);
                }
            }

            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(implicitWait);

            return Driver;
        }

        // Connect and auto-update chromedriver.exe
        public static ChromeDriver ConnectUpdate(string chromedriverPath, string chromeBinPath = null, string defDwdDir = @"%USERPROFILE%\Downloads", int implicitWait = 500)
        {
            Action connect = () =>
            {
                Driver = Connect(chromedriverPath, chromeBinPath, defDwdDir, implicitWait);
            };

            try
            {
                connect();
            }
            catch (ChromeDriverDeprecatedException)
            {
                UpdateChromeDriver(chromedriverPath, chromeBinPath);
                connect();
            }

            return Driver;
        }

        public static string GetDriverVersion()
        {
            return (Driver.Capabilities.GetCapability("chrome") as Dictionary<string, object>)["chromedriverVersion"].ToString();
        }

        public static void SwitchToTab(IWebDriver driver, string tabName)
        {
            string oldTab = driver.Title;
            string newTab;
            bool found = false;

            foreach (var handle in driver.WindowHandles)
            {

                driver.SwitchTo().Window(handle);
                if (Regex.IsMatch(driver.Title, tabName, RegexOptions.IgnoreCase))
                {
                    found = true;
                    break;
                }
            }

            if (found == false)
            {
                throw new TabNotFoundException(tabName);
            }
            newTab = driver.Title;
            Log.Write(string.Format("Tab motion: \"{0}\" -> \"{1}\"", oldTab, newTab));
        }

        public static void SetLinkOpenOnSameTab(IWebDriver driver, string xpathSelector)
        {
            string funcDef = "let getElementByXpath = " +
                "(p) => document.evaluate(p, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;";
            //string jsSetToSelf = @"getElementByXpath(""{0}"").setAttribute(""target"", ""self"");";
            string jsSetToSelf = @"getElementByXpath(""{0}"").target = 'self'";

            string cmd = funcDef + "\n" + string.Format(jsSetToSelf, xpathSelector);

            driver.ExecuteJavaScript(cmd);

        }

        public static void UpdateChromeDriver(string installPathDir, string chromeBinPath = null) =>
            Task.WaitAll(UpdateChromeDriverAsync(installPathDir, chromeBinPath));

        public static async Task UpdateChromeDriverAsync(string installPathDir, string chromeBinPath = null)
        {
            //Rpa.MessageBox("Starting to update ChromeDriver");
            Log.Write("Starting to update ChromeDriver...");

            // File names and paths
            string exeFileName = "chromedriver.exe";
            string zipFileName = "chromedriver_win32.zip";
            string savePathDir = Ut.GetFullPath(@"%USERPROFILE%\Downloads\SeleniumChromeDriver");
            string savePathZip = Path.Combine(savePathDir, zipFileName);
            string exeFullPathDownloads = Path.Combine(savePathDir, exeFileName);
            string exeFullPathInstall = Path.Combine(installPathDir, exeFileName);

            string chromeMajorVersion = GetChromeMajorVersion(chromeBinPath);

            // mount the url with the major version
            string url = $"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{chromeMajorVersion}";
            // api call to discover the ChromeDriver version compatible with system's Chrome
            string chromeDriverVersion = await Ut.MakeApiCallAsync(url);

            // mount the url with the ChromeDriver full version spec
            string chromeDriverUrl = $"https://chromedriver.storage.googleapis.com/{chromeDriverVersion}/{zipFileName}";

            // create download destination dir if it doesn't exist
            if (!Directory.Exists(savePathDir))
            {
                Directory.CreateDirectory(savePathDir);
            }

            // download the zip
            string dwdResult = await Ut.MakeApiCallAsync(chromeDriverUrl, saveAs: savePathZip);

            // delete the files because Unzip can't overwrite
            foreach (string file in Ut.MatchAllFiles(@"chromedriver\.exe|LICENSE\.chromedriver", savePathDir))
                File.Delete(file);
            // now we guarantee that Unzip() will always occur with success
            Ut.Unzip(savePathZip, savePathDir);

            string oldVersion = Ut.RunPromptCommand(Bin, "--version");
            //string oldVersion = GetDriverVersion();
            // move to the folder listed in PATH to actually install (suggestion: %APPDATA%\Selenium\Drivers)
            Ut.MoveFileForced(exeFullPathDownloads, exeFullPathInstall);
            string newVersion = Ut.RunPromptCommand(Bin, "--version");
            //string newVersion = chromeDriverVersion;

            Log.Write(string.Join(Environment.NewLine,
                $"ChromeDriver updated.",
                $"  Old version: {oldVersion}",
                $"  New version: {newVersion}"));
        }

        public static string GetChromeMajorVersion(string chromeBinPath = null)
        {
            // Get the major version number from the currently installed chrome in system
            string cmdResult;
            string chromeVersion;
            if (chromeBinPath == null)
            {
                // if bin path is NOT given, then check the chrome system installed version through windows register
                cmdResult = Ut.RunPromptCommand("reg", @"query ""HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon"" /v version");
                chromeVersion = Ut.Match(cmdResult, @"(?<=version[\s\w]+)\d+\.\d+\.\d+(\.\d+)?");
            }
            else
            {
                // if Chrome binary path is given, then check the Chrome version using this file
                cmdResult = Ut.RunPromptCommand("wmic", $"datafile where name=\"{Ut.GetFullPath(chromeBinPath).Replace(@"\", @"\\")}\" get Version /value");
                chromeVersion = Ut.Match(cmdResult, @"(?<=Version=)[\w.]+");
            }

            string chromeMajorVersion = Ut.Match(chromeVersion, @"^\d+(?=\.)");

            return chromeMajorVersion;
        }
    }
}
