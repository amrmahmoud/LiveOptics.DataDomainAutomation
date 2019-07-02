﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NUnit.Framework;
using LiveOptics.DataDomainAutomation;
using System.Diagnostics;

namespace LiveOptics.DataDomainAutomation
{
    [TestFixture]
    public class ExcelTests
    {
        private readonly ExcelComparator _excelCompare = new ExcelComparator();
        string location;

        [OneTimeSetUp]
        public void TestInitialize()
        {
            location = Environment.CurrentDirectory;
            var pathToExecutable = "cd " + location + "\\..\\..\\..\\Executable";
            var executableCommand = "DDLS.exe -inputFolder \"..\\ASUPS\" -outputFolder \"..\\ASUPS\"";
            //Running cmd and passing arguments
            Process cmd = new Process();
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.Start();
            //Sending command to run .exe on asup files to generate xlxs
            cmd.StandardInput.WriteLine(pathToExecutable);
            cmd.StandardInput.WriteLine(executableCommand);
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();
            Console.WriteLine(cmd.StandardOutput.ReadToEnd());
        }
        [Test]
        public void VerifyXlsxDDConfigInfo()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ASUPS");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + "Config" + "*", SearchOption.AllDirectories);
            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                Console.WriteLine(fullName);
                string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullName));
                Console.WriteLine(parentDirectoryName);

                string actualExcel = fullName;
                string expectedExcel = "..\\..\\..\\XLSX Templates\\Data Domain Config " +parentDirectoryName+".xlsx";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }
        [Test]
        public void VerifyXlsxDDHistory()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ASUPS");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + "History" + "*", SearchOption.AllDirectories);

            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                Console.WriteLine(fullName);
                string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullName));
                Console.WriteLine(parentDirectoryName);

                string actualExcel = fullName;
                var expectedExcel = "..\\..\\..\\XLSX Templates\\Data Domain History " +parentDirectoryName+ ".xlsx";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }
        [Test]
        public void VerifyXlsxDDReplicationMapping()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ASUPS");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + "Replication" + "*", SearchOption.AllDirectories);

            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                Console.WriteLine(fullName);
                string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullName));
                Console.WriteLine(parentDirectoryName);

                string actualExcel = fullName;
                var expectedExcel = "..\\..\\..\\XLSX Templates\\Data Domain Replication Mapping " +parentDirectoryName+ ".xlsx";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }
        [Test]
        public void VerifyXlsxDDBoostInfo()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ASUPS");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + "BOOST Information" + "*", SearchOption.AllDirectories);

            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                Console.WriteLine(fullName);
                string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullName));
                Console.WriteLine(parentDirectoryName);

                string actualExcel = fullName;
                var expectedExcel = "..\\..\\..\\XLSX Templates\\Data Domain BOOST Information " +parentDirectoryName+ ".xlsx";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }

    }
}