using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LiveOptics.DataDomainAutomation
{
    public class ExcelComparator1
    {
        public ComparisonResponseModel AreEqual(string pathToExpectedFile, string pathToActualFile,
           bool compareHeadingsOnly = false)
        {
            ComparisonResponseModel result = new ComparisonResponseModel();
            string responseText;
            bool passing = true;

            using (SpreadsheetDocument expectedOpenExcel = SpreadsheetDocument.Open(pathToExpectedFile, false))
            {
                using (SpreadsheetDocument actualOpenExcel = SpreadsheetDocument.Open(pathToActualFile, false))
                {
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------");
                    Console.WriteLine($"Beginning Excel comparisson...");
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------");
                    WorkbookPart actualWorkbookPart = actualOpenExcel.WorkbookPart;
                    IEnumerable<WorksheetPart> actualWorksheetParts = actualWorkbookPart.WorksheetParts;
                    WorkbookPart expectedWorkbookPart = expectedOpenExcel.WorkbookPart;
                    IEnumerable<WorksheetPart> expectedWorksheetParts = expectedWorkbookPart.WorksheetParts;

                    if (actualWorkbookPart.Workbook.Sheets.Count() != expectedWorkbookPart.Workbook.Sheets.Count())
                    {
                        responseText = "Excel Comparator failed when comparing sheet count on both workbooks.";
                        Console.WriteLine(responseText);
                        passing = false;
                    }
                    else Console.WriteLine("Sheets Quantity Match...");

                    foreach (Sheet sheet in actualWorkbookPart.Workbook.Sheets)
                    {
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------");
                        Console.WriteLine($"Sheet Under Test: {sheet.Name}");
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------");

                        Sheet expectedSheet = expectedWorkbookPart.Workbook.Descendants<Sheet>().Where(sht => sht.Name == sheet.Name.InnerText).FirstOrDefault();

                        if (String.IsNullOrEmpty(expectedSheet.Name))
                        {
                            responseText =
                                $"Excel Comparator did not find {sheet.Name} in expected Excel";
                            Console.WriteLine(responseText);
                            passing = false;
                        }
                        else Console.WriteLine("Sheet Names Match...");

                        WorksheetPart actualWorksheetPart = (WorksheetPart)actualWorkbookPart.GetPartById(sheet.Id.Value);
                        WorksheetPart expectedWorksheetPart = (WorksheetPart)expectedWorkbookPart.GetPartById(expectedSheet.Id.Value);

                        //Column usage check
                        int usedColumnCountInActual = actualWorksheetPart.Worksheet.Descendants<Column>().Count();
                        int usedColumnCountInExpected = expectedWorksheetPart.Worksheet.Descendants<Column>().Count();
                        if (usedColumnCountInActual != usedColumnCountInExpected)
                        {
                            responseText =
                                $"Excel Comparator failed when comparing count of columns used within sheet {sheet.Name} on both workbooks. Expected {usedColumnCountInExpected} but actual value was {usedColumnCountInActual}";
                            Console.WriteLine(responseText);
                            passing = false;
                        }
                        else Console.WriteLine("Used Column Quantity Match...");

                        SheetData actualSheetData = actualWorksheetPart.Worksheet.Elements<SheetData>().First();
                        SheetData expectedSheetData = expectedWorksheetPart.Worksheet.Elements<SheetData>().First();
                        //Row usage check
                        int usedRowCountInActual = actualSheetData.Elements<Row>().Count();
                        int usedRowCountInExpected = expectedSheetData.Elements<Row>().Count();
                        if (usedRowCountInActual != usedRowCountInExpected)
                        {
                            responseText =
                                $"Excel Comparator failed when comparing count of rows used within sheet {sheet.Name} on both workbooks. Expected {usedRowCountInExpected} but actual value was {usedRowCountInActual}";
                            Console.WriteLine(responseText);
                            passing = false;
                        }
                        else Console.WriteLine("Used Row Quantity Match...");
                        var actualRows = actualSheetData.ChildElements;
                        var expectedRows = expectedSheetData.ChildElements;
                        for(var i = 0; i <actualRows.Count; i++)
                        {
                            var actualCells = actualRows[i].ChildElements;
                            var expectedCells = expectedRows[i].ChildElements;

                            if (actualCells.Count != expectedCells.Count)
                            {
                                responseText =
                                    $"Excel Comparator failed when comparing count of cells used within sheet {sheet.Name} on both workbooks. Expected {expectedCells.Count} but actual value was {actualCells.Count}";
                                Console.WriteLine(responseText);
                                passing = false;
                            }

                            for (int j = 0; j <actualCells.Count; j++)
                            {
                                string actualCellValue = actualCells[j]?.InnerText?.ToString();
                                string expectedCellValue = expectedCells[j]?.InnerText?.ToString();

                                if (!string.Equals(actualCellValue, expectedCellValue))
                                {
                                    responseText =
                                        $"Excel Comparator failed when comparing value of cells used within sheet {sheet.Name}, row {i+1}, cell {j+1} on both workbooks. Expected {expectedCellValue} but actual value was {actualCellValue}";
                                    Console.WriteLine(responseText);
                                    passing = false;
                                }
                            }
                        }
                        Console.WriteLine($"End of Cell Checks for Sheet: {sheet.Name}");
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------");
                    }
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------");
                    Console.WriteLine($"End of Excel comparisson...");
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------");
                    if (passing == true)
                    {
                        result.Passed = true;
                        result.ResponseText = "";
                        return result;
                    }
                    else
                    {
                        responseText = $"EXCEL COMPARATOR FAILED";
                        Console.WriteLine(responseText);
                        result.Passed = false;
                        result.ResponseText = responseText;

                        return result;
                    }
                }
            }
        }
    }
}
