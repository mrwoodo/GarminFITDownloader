using Newtonsoft.Json.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;

namespace GarminFITDownloader
{
    class Program
    {
        /// <summary>
        /// Utility to download all your Fit files from the Garmin Portal
        /// Uses screenscraping to login / fetch activity list
        /// Works as at 26 Jan 2019
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            using (var driver = new ChromeDriver())
            {
                //Settings
                var username = "{{{your username (email}}}";
                var password = "{{{your password}}}";
                var destFolder = @"c:\temp\";

                //Login to Garmin Portal
                driver.Navigate().GoToUrl("https://connect.garmin.com/en-US/signin");
                driver.SwitchTo().DefaultContent();

                var containerFrame = driver.FindElement(By.Id("gauth-widget-frame-gauth-widget"));
                driver.SwitchTo().Frame(containerFrame);

                var userNameField = driver.FindElementById("username");
                var userPasswordField = driver.FindElementById("password");
                var loginButton = driver.FindElementById("login-btn-signin");

                userNameField.SendKeys(username);
                userPasswordField.SendKeys(password);
                loginButton.Click();

                //Move to Activities page
                driver.Navigate().GoToUrl("https://connect.garmin.com/modern/activities");

                //Now fetch json list of activities
                //TODO: I only have 100 activities so not sure what's the limit you can fetch in a single request - Set to 500 in this example
                driver.Navigate().GoToUrl("https://connect.garmin.com/modern/proxy/activitylist-service/activities/search/activities?limit=500&start=0");
                var json = driver.FindElementByTagName("body");
                dynamic activities = JArray.Parse(json.Text);

                //Here we use a webclient to download each Activity zip file, so we have to pass
                //in the Cookies from the Selenium session to the WebClient
                using (var client = new WebClient())
                {
                    var cookies = driver.Manage().Cookies.AllCookies;
                    client.Headers[HttpRequestHeader.Cookie] = string.Join("; ", cookies.Select(c => string.Format("{0}={1}", c.Name, c.Value)));

                    foreach (var activity in activities)
                    {
                        //Get an activity zip file
                        var url = $"https://connect.garmin.com/modern/proxy/download-service/files/activity/{activity.activityId}";
                        var filename = $"{activity.startTimeLocal}-{activity.activityName}-{activity.activityId}".Replace(" ", "").Replace(":", "");
                        client.DownloadFile(url, $"{destFolder}{filename}.zip");

                        //Extract the fit file from it
                        using (ZipArchive archive = ZipFile.OpenRead($"{destFolder}{filename}.zip"))
                        {
                            var fitFile = archive.Entries[0];
                            fitFile.ExtractToFile($"{destFolder}{filename}.fit", true);

                            //Set last-modified of the fit file to the time the activity started
                            var fileTime = DateTime.ParseExact(activity.startTimeLocal.ToString(), "yyyy-MM-dd HH:mm:ss", null);
                            File.SetLastWriteTime($"{destFolder}{filename}.fit", fileTime);
                        }
                    }
                }

                driver.Close();
            }
        }
    }
}