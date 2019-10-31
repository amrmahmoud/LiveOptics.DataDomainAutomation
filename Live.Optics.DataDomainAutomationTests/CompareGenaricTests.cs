using System;
using System.IO;
using NUnit.Framework;
using LiveOptics.DataDomainAutomation;
using System.Diagnostics;

namespace Live.Optics.DataDomainAutomation.Tests
{
    [TestFixture]
    public class CompareGenaricTests
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
        public void VerifyXlsxDDVtV()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ASUPS");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + "VtV" + "*", SearchOption.AllDirectories);
            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                Console.WriteLine(fullName);
                string parentDirectoryName = Path.GetFileName(Path.GetDirectoryName(fullName));
                Console.WriteLine(parentDirectoryName);

                string actualExcel = fullName;
                string expectedExcel = "..\\..\\..\\XLSX Templates\\Data Domain VtV " + parentDirectoryName + ".xlsx";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }
    }
}
