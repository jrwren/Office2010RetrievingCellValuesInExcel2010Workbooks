using System;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace RetrieveCellValue
{
  class Program
  {
    static void Main(string[] args)
    {
      // Place a double value into A1, a date into A2, a string into A3, and a TRUE value into A4.
      const string fileName = @"C:\temp\GetCellValue.xlsx";
      string value = XLGetCellValue(fileName, "Sheet1", "A1");
      Console.WriteLine(value);
      value = XLGetCellValue(fileName, "Sheet1", "A2");
      Console.WriteLine(value);
      value = XLGetCellValue(fileName, "Sheet1", "A3");
      Console.WriteLine(value);
      value = XLGetCellValue(fileName, "Sheet1", "A4");
      Console.WriteLine(value);
    }

    // Get the value of a cell, given a file name, sheet name, and address name.
    public static string XLGetCellValue(string fileName, string sheetName, string addressName)
    {
      string value = null;

      using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
      {
        WorkbookPart wbPart = document.WorkbookPart;

        // Find the sheet with the supplied name, and then use that Sheet object
        // to retrieve a reference to the appropriate worksheet.
        Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
          Where(s => s.Name == sheetName).FirstOrDefault();

        if (theSheet == null)
        {
          throw new ArgumentException("sheetName");
        }

        // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
        // a reference to the cell whose address matches the address you've supplied:
        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
        Cell theCell = wsPart.Worksheet.Descendants<Cell>().
          Where(c => c.CellReference == addressName).FirstOrDefault();

        // If the cell doesn't exist, return an empty string:
        if (theCell != null)
        {
          value = theCell.InnerText;

          // If the cell represents an integer number, you're done. 
          // For dates, this code returns the serialized value that 
          // represents the date. The code handles strings and booleans
          // individually. For shared strings, the code looks up the corresponding
          // value in the shared string table. For booleans, the code converts 
          // the value into t he words TRUE or FALSE.
          if (theCell.DataType != null)
          {
            switch (theCell.DataType.Value)
            {
              case CellValues.SharedString:
                // For shared strings, look up the value in the shared strings table.
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                // If the shared string table is missing, something's wrong.
                // Just return the index that you found in the cell.
                // Otherwise, look up the correct text in the table.
                if (stringTable != null)
                {
                  value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
                break;

              case CellValues.Boolean:
                switch (value)
                {
                  case "0":
                    value = "FALSE";
                    break;
                  default:
                    value = "TRUE";
                    break;
                }
                break;
            }
          }
        }
      }
      return value;
    }

  }
}
