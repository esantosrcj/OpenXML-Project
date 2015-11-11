using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

using System.Text.RegularExpressions;

namespace OpenXmltoExcel
{
    public class CopyXlsx
    {
        private const uint ROW = 15;
        private const uint COL = 10;

        // Create a Spreadsheet document
        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part
        private static void CreateParts(SpreadsheetDocument document)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            // Add a WorkbookPart to the document.
            WorkbookPart starterWorkbookPart = document.AddWorkbookPart();
            starterWorkbookPart.Workbook = new Workbook();
            starterWorkbookPart.Workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart starterWorksheetPart = starterWorkbookPart.AddNewPart<WorksheetPart>();

            // SheetData: Represents a cell table. Expresses information about each cell, grouped together by rows in the worksheet.
            starterWorksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets firstSheets = document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet leadingSheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(starterWorksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            firstSheets.Append(leadingSheet);

            // Save the document
            starterWorkbookPart.Workbook.Save();

            // Close the document
            document.Close();
        }

        // Given a WorkbookPart, inserts a new worksheet.
        public static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            // GetFirstChild <T>: Find the first child element in type T
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Add another worksheet to document.
        public void AddWorksheetToDocument(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                WorksheetPart anotherWorksheetPart = InsertWorksheet(document.WorkbookPart);
            }
        }

        // Insert data values into sheet data.
        private static void InsertValuesInSheetData(WorksheetPart worksheetPart, List<string> cellAddresses, List<string> values)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            int indexOfValue = 0;

            foreach (var item in cellAddresses)
            {
                Row row;
                uint rowIndex = (uint)int.Parse(Regex.Match(item, @"\d+").Value);

                // If the worksheet does not contain a row with the specified row index, insert one.
                if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
                {
                    row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
                }
                else
                {
                    row = new Row() { RowIndex = rowIndex };
                    sheetData.Append(row);
                }

                // If there is not a cell with the specified column name, insert one.
                if (row.Elements<Cell>().Where(c => c.CellReference.Value == item).Count() > 0)
                {
                    var insertCell = row.Elements<Cell>().Where(c => c.CellReference.Value == item).First();
                    insertCell = new Cell
                    {
                        CellReference = item,
                        CellValue = new CellValue(values[indexOfValue]),
                        DataType = new EnumValue<CellValues>(CellValues.String)
                    };
                    worksheetPart.Worksheet.Save();
                }
                else
                {
                    Cell refCell = null;
                    foreach (Cell cellInRow in row.Elements<Cell>())
                    {
                        if (string.Compare(cellInRow.CellReference.Value, item, true) > 0)
                        {
                            refCell = cellInRow;
                            break;
                        }
                    }

                    var newCell = new Cell()
                    {
                        CellReference = item,
                        CellValue = new CellValue(values[indexOfValue]),
                        DataType = new EnumValue<CellValues>(CellValues.String)
                    };
                    row.InsertBefore(newCell, refCell);
                    worksheetPart.Worksheet.Save();
                }

                // Increment the index of the value.
                indexOfValue++;
            }
        }

        // Add the data to the worksheet.
        public void AddDataToWorksheet(string fileName)
        {
            // Generate the data to insert into cells.
            //string[,] cellStrings = TestStrings(ROW, COL);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                // Get a list of worksheet parts.
                var worksheetParts = document.WorkbookPart.GetPartsOfType<WorksheetPart>();

                //foreach (var wrksheetPart in worksheetParts)
                //{
                //    InsertValuesInSheetData(wrksheetPart, ROW, COL, cellStrings);
                //}

                // Close the document
                document.Close();
            }
        }

        // Add copied data to worksheet.
        public void AddDataToWorksheet(string fileName, string sheetName, List<string> cellAddresses, List<string> copiedData)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                // Get a list of worksheet parts.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("Sheet name does not exist.");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                InsertValuesInSheetData(wsPart, cellAddresses, copiedData);

                // Close the document
                document.Close();
            }
        }

        // Retrieve the value of a cell, given a file name, sheet name, and address name.
        public static string GetCellValue(string fileName, string sheetName, string cellAddress)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("Sheet name does not exist.");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellAddress).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
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

        // Read the spreadsheet and return with values.
        public static List<string> ReadSpreadsheet(string fileName, string sheetName)
        {
            var fileData = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false)) // read only
            {
                // Get sheet data.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("Sheet name does not exist.");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                foreach (Cell cell in wsPart.Worksheet.Descendants<Cell>())
                {
                    fileData.Add(GetCellValue(fileName, sheetName, (cell.CellReference.Value.ToString())));
                }

                // Close document.
                document.Close();
            }

            return fileData;
        }

        // Get the number of sheets in a spreadsheet.
        public static int GetNumberOfSheets(string fileName)
        {
            int numberOfSheet = 0;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                var sheet = wbPart.Workbook.Descendants<Sheet>();

                // Throw an exception if there is no sheet.
                if (sheet == null)
                {
                    throw new ArgumentException("Sheet name does not exist.");
                }

                foreach (var item in sheet)
                {
                    numberOfSheet++;
                }

                // Close document.
                document.Close();
            }
            return numberOfSheet;
        }

        // Get the cell references of the existing file.
        public static List<string> GetCellReferences(string fileName, string sheetName)
        {
            var cellReferences = new List<string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = doc.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("Sheet name does not exist.");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                foreach (Cell cell in wsPart.Worksheet.Descendants<Cell>())
                {
                    cellReferences.Add(cell.CellReference.Value.ToString());
                }

                // Close document.
                doc.Close();
            }

            return cellReferences;
        }

        // Copy the data from existing file and save into new file.
        public void CopyAndSave(string newFile, string existingFile)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(newFile, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }

            // Get the number of sheets from the existing file
            int numOfSheets = GetNumberOfSheets(existingFile);

            // There should already be one sheet in the new file.
            if (numOfSheets > 1)
            {
                for (int i = 1; i < numOfSheets; i++)
                {
                    // Add worksheets to new document.
                    AddWorksheetToDocument(newFile);
                }
            }

            string sheetName = "";
            for (int i = 1; i <= numOfSheets; i++)
            {
                sheetName = "Sheet" + i;
                var cells = GetCellReferences(existingFile, sheetName);
                var listOfValues = ReadSpreadsheet(existingFile, sheetName);

                AddDataToWorksheet(newFile, sheetName, cells, listOfValues);
            }
        }
    }
}
