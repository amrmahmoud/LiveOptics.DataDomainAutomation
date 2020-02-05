using System;
using System.IO;
using NUnit.Framework;
using LiveOptics.DataDomainAutomation;
using System.Diagnostics;


namespace Live.Optics.DataDomainAutomation.Tests
{
    [TestFixture]
    public class BasicComparator
    {
        private readonly ExcelComparator _excelCompare = new ExcelComparator();

        [Test]
        public void VerifyXlsx2DDVtV()
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
