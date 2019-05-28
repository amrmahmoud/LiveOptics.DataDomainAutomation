using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using LiveOptics.DataDomainAutomation;

namespace LiveOptics.DataDomainAutomation
{
    public class ExcelComparator
    {
        
        public ComparisonResponseModel AreEqual( string pathToExpectedFile, string pathToActualFile,
            bool compareHeadingsOnly = false )
        {
            ComparisonResponseModel result = new ComparisonResponseModel();
            string responseText;
            bool passing = true;

            //Setup excel instance and Open Workbooks
            //Closed xml Varient
            using ( XLWorkbook expectedClosedExcel = new XLWorkbook( pathToExpectedFile ) )
            {
                using ( XLWorkbook actualClosedExcel = new XLWorkbook( pathToActualFile ) )
                {
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------" );
                    Console.WriteLine( $"Beginning Excel comparisson..." );
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------" );

                    //First compare the number of sheets in the workbook. if these differ, fail as there is a obvious difference at that stage
                    if ( actualClosedExcel.Worksheets.Count != expectedClosedExcel.Worksheets.Count )
                    {
                        responseText = "Excel Comparator failed when comparing sheet count on both workbooks.";
                        Console.WriteLine( responseText );
                        passing = false;
                        //result.Passed = false;
                        //result.ResponseText = responseText;
                        //return result;
                    }
                    else Console.WriteLine( "Sheets Quantity Match..." );

                    //If above passes, lets go sheet by sheet and compare name and usage metrics
                    foreach ( IXLWorksheet sheet in actualClosedExcel.Worksheets )
                    {
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------" );
                        Console.WriteLine( $"Sheet Under Test: {sheet.Name}" );
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------" );
                        //create an instance of the expected worksheet using the same index of the actual worksheet
                        IXLWorksheet expectedWorksheet = expectedClosedExcel.Worksheets.Worksheet( sheet.Position );

                        //sheet name check
                        if ( sheet.Name != expectedWorksheet.Name )
                        {
                            responseText =
                                $"Excel Comparator failed when comparing sheet names on both workbooks. Expected {expectedWorksheet.Name} but actual value was {sheet.Name}";
                            Console.WriteLine( responseText );
                            passing = false;
                            //result.ResponseText = responseText;
                            //result.Passed = false;

                            //return result;
                        }
                        else Console.WriteLine( "Sheet Names Match..." );

                        //Column usage check
                        int usedColumnCountInActual = sheet.ColumnsUsed().Count();
                        int usedColumnCountInExpected = expectedWorksheet.ColumnsUsed().Count();
                        if ( usedColumnCountInActual != usedColumnCountInExpected )
                        {
                            responseText =
                                $"Excel Comparator failed when comparing count of columns used within sheet {sheet.Name} on both workbooks. Expected {usedColumnCountInExpected} but actual value was {usedColumnCountInActual}";
                            Console.WriteLine( responseText );
                            passing = false;
                            //result.ResponseText = responseText;
                            //result.Passed = false;

                            //return result;
                        }
                        else Console.WriteLine( "Used Column Quantity Match..." );

                        //Heading Values check
                        int currentColIndex = 1;
                        int rowIndex = 1;
                        while ( currentColIndex <= usedColumnCountInActual )
                        {
                            if ( sheet.Cell( rowIndex, currentColIndex ).Value.ToString() != expectedWorksheet
                                     .Cell( rowIndex, currentColIndex ).Value.ToString() )
                            {
                                responseText =
                                    $"Excel Comparator failed when comparing the column headings within sheet {sheet.Name} on both workbooks. Expected {expectedWorksheet.Cell( rowIndex, currentColIndex ).Value} but actual value was {sheet.Cell( rowIndex, currentColIndex ).Value}";
                                Console.WriteLine( responseText );
                                passing = false;
                                //result.ResponseText = responseText;
                                //result.Passed = false;

                                //return result;
                            }
                            Console.WriteLine( sheet.Cell( rowIndex, currentColIndex ).Value );
                            currentColIndex = currentColIndex + 1;
                        }
                        Console.WriteLine( "Sheet Headings Match..." );

                        if ( !compareHeadingsOnly )
                        {
                            //Row usage check
                            int usedRowCountInActual = sheet.RowsUsed().Count();
                            int usedRowCountInExpected = expectedWorksheet.RowsUsed().Count();
                            if ( usedRowCountInActual != usedRowCountInExpected )
                            {
                                responseText =
                                    $"Excel Comparator failed when comparing count of rows used within sheet {sheet.Name} on both workbooks. Expected {usedRowCountInExpected} but actual value was {usedRowCountInActual}";
                                Console.WriteLine( responseText );
                                passing = false;
                                //result.ResponseText = responseText;
                                //result.Passed = false;

                                //return result;
                            }
                            else Console.WriteLine( "Used Row Quantity Match..." );

                            //Cell usage check
                            int usedCellCountInActual = sheet.CellsUsed().Count();
                            int usedCellCountInExpected = expectedWorksheet.CellsUsed().Count();
                            if ( usedCellCountInActual != usedCellCountInExpected )
                            {
                                responseText =
                                    $"Excel Comparator failed when comparing count of cells used within sheet {sheet.Name} on both workbooks. Expected {usedCellCountInExpected} but actual value was {usedCellCountInActual}";
                                Console.WriteLine( responseText );
                                passing = false;
                                //result.ResponseText = responseText;
                                //result.Passed = false;

                                //return result;
                            }
                            else Console.WriteLine( "Used Cell Quantity Match..." );

                            //After basic value checks pass, lets go a level deeper and compare cell content for each populated cell.
                            foreach ( IXLCell cell in sheet.CellsUsed() )
                            {
                                int cellRowUnderTest = cell.WorksheetRow().RowNumber();
                                int cellColumnUnderTest = cell.WorksheetColumn().ColumnNumber();
                                Console.WriteLine(
                                    "--------------------------------------------------------------------------------" );
                                Console.WriteLine(
                                    $"Cell Under Test: Cell Row Index {cellRowUnderTest}, Cell Column Index {cellColumnUnderTest}" );
                                Console.WriteLine(
                                    "--------------------------------------------------------------------------------" );
                                //Create an instance of the expected cell value to use. This is based on the current cell in actual file.
                                IXLCell expectedCell = expectedWorksheet.Cell( cellRowUnderTest, cellColumnUnderTest );

                                //Cell value check
                                string actualCellValue = cell.Value.ToString();
                                string expectedCellValue = expectedCell.Value.ToString();

                                if ( !string.Equals( actualCellValue, expectedCellValue ) )
                                {
                                    responseText =
                                        $"Excel Comparator failed when comparing value of cells used within sheet {sheet.Name} on both workbooks. Expected {expectedCellValue} but actual value was {actualCellValue}";
                                    Console.WriteLine( responseText );
                                    passing = false;
                                    //result.ResponseText = responseText;
                                    //result.Passed = false;

                                    //return result;
                                }
                               else Console.WriteLine( "Cell Contents Match..." );
                                    Console.WriteLine(
                                    "--------------------------------------------------------------------------------" );
                            }
                        }
                        Console.WriteLine( $"End of Cell Checks for Sheet: {sheet.Name}" );
                        Console.WriteLine(
                            "--------------------------------------------------------------------------------" );
                    }
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------" );
                    Console.WriteLine( $"End of Excel comparisson..." );
                    Console.WriteLine(
                        "--------------------------------------------------------------------------------" );
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