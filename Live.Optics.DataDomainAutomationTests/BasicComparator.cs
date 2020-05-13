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
        private readonly ExcelComparator1 _excelCompare = new ExcelComparator1();

        [Test]
        public void VerifyXlsx2DDVtV()
        {
            DirectoryInfo hdDirectoryInWhichToSearch = new DirectoryInfo(@"..\..\..\ActualExcel");
            FileSystemInfo[] filesAndDirs = hdDirectoryInWhichToSearch.GetFiles("*" + ".xlsx" + "*", SearchOption.AllDirectories);
            foreach (FileSystemInfo foundFile in filesAndDirs)
            {
                string fullName = foundFile.FullName;
                string filename = Path.GetFileName(fullName);

                string actualExcel = fullName;
                string expectedExcel = $"..\\..\\..\\ExpectedExcel\\{filename}";

                //compare files
                ComparisonResponseModel fileequalitycheck = _excelCompare.AreEqual(expectedExcel, actualExcel);
                Assert.IsTrue(fileequalitycheck.Passed);
                Assert.IsEmpty(fileequalitycheck.ResponseText);
            }
        }

    }
}
