﻿using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Configuration;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Z01_NU_nCore.Tests
{
    [SetUpFixture]
    public class GlobalTestSetup
    {
        public static ExtentReports extent;
        public static string reportPath;
        Process driverProcess=null;

        [OneTimeSetUp]
        public void GlobalSetup()
        {
            try
            {
                driverProcess = Process.Start(@"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe");
            }
            catch {
                Console.WriteLine("seems driver is already started");
            }
            string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            // string projectPath = @"C:\Program Files (x86)\Jenkins\workspace\TestRunner";
            reportPath = assemblyPath + "..\\..\\Reports\\TestRunReport.html";
            extent = new ExtentReports();
            ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter(reportPath);
            htmlReporter.Config.Theme = Theme.Standard;
            htmlReporter.Config.DocumentTitle = "Zero Windows";
            htmlReporter.Config.ReportName = "Zero";
            extent.AttachReporter(htmlReporter);

        }

        [OneTimeTearDown]
        public void GlobalTeardown()
        {
            extent.Flush();
            if (driverProcess != null)
            {
                driverProcess.Kill();
            }
        }
    }
}