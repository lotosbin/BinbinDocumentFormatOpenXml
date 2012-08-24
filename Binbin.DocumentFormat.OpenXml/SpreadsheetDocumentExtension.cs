#region

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

#endregion

namespace Binbin.DocumentFormat.OpenXml
{
    public static class SpreadsheetDocumentExtension
    {
        internal static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (!sheets.Any())
            {
                // The specified worksheet does not exist.
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            var worksheetPart = (WorksheetPart) document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }


        // inserts a new worksheet and writes the text to cell "A1" of the new worksheet.
        /// <summary>
        ///   Given a document name and text,
        /// </summary>
        /// <param name="docName"> </param>
        /// <param name="text"> </param>
        internal static void InsertText(string docName, string text)
        {
            // Open the document for editing.
            using (var spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                // Insert the text into the SharedStringTablePart.
                int index = InsertSharedStringItem(text, shareStringPart);

                // Insert a new worksheet.
                var worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);

                // Insert cell A1 into the new worksheet.
                var cell = GetCell(worksheetPart, 1, "A");

                // Set the value of cell A1.
                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                // Save the new worksheet.
                //worksheetPart.Worksheet.Save();
            }
        }

        /// <summary>
        ///   Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        ///   and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text"> </param>
        /// <param name="shareStringPart"> </param>
        /// <returns> </returns>
        internal static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            //shareStringPart.SharedStringTable.Save();

            return i;
        }

        /// <summary>
        ///   Given a document name, inserts a new worksheet.
        /// </summary>
        /// <param name="docName"> </param>
        internal static void InsertWorksheet(string docName)
        {
            // Open the document for editing.
            using (var spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                // Add a blank WorksheetPart.
                var newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                var relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

                // Get a unique ID for the new worksheet.
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Any())
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                // Give the new worksheet a name.
                var sheetName = "Sheet" + sheetId;

                // Append the new worksheet and associate it with the workbook.
                var sheet = new Sheet() {
                        Id = relationshipId,
                        SheetId = sheetId,
                        Name = sheetName
                };
                sheets.Append(sheet);
            }
        }

        /// <summary>
        ///   Given a WorkbookPart, inserts a new worksheet.
        /// </summary>
        /// <param name="workbookPart"> </param>
        /// <returns> </returns>
        internal static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            //newWorksheetPart.Worksheet.Save();

            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            var sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            var sheet = new Sheet() {
                    Id = relationshipId,
                    SheetId = sheetId,
                    Name = sheetName
            };
            sheets.Append(sheet);
            //workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        /// <summary>
        ///   Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellName"> </param>
        /// <returns> </returns>
        internal static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);
            return match.Value;
        }

        /// <summary>
        ///   Given a cell name, parses the specified cell to get the row index.
        /// </summary>
        /// <param name="cellName"> </param>
        /// <returns> </returns>
        internal static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);
            return uint.Parse(match.Value);
        }

        #region get cell

        /// <summary>
        ///   Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        ///   If the cell already exists, returns it.
        /// </summary>
        /// <param name="worksheet"> </param>
        /// <param name="rowIndex"> </param>
        /// <param name="columnName"> </param>
        /// <returns> </returns>
        internal static Cell GetCell(Worksheet worksheet, uint rowIndex, string columnName)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var row = GetRow(sheetData, rowIndex);

            string cellReference = columnName + rowIndex;
            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Any())
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            var refCell = row.Elements<Cell>().FirstOrDefault(cell => cell.CellReference.Value == cellReference);

            var newCell = new Cell() {
                    CellReference = cellReference
            };
            row.InsertBefore(newCell, refCell);

            //worksheet.Save();
            return newCell;
        }

        internal static Cell GetCell(WorksheetPart worksheetPart, uint rowIndex, string columnName)
        {
            return GetCell(worksheetPart.Worksheet, rowIndex, columnName);
        }

        #endregion

        #region get row

        /// <summary>
        ///   If the worksheet does not contain a row with the specified row index, insert one.
        /// </summary>
        /// <param name="sheetData"> </param>
        /// <param name="rowIndex"> </param>
        /// <returns> </returns>
        internal static Row GetRow(SheetData sheetData, uint rowIndex)
        {
            // If the worksheet does not contain a row with the specified row index, insert one.
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Any())
            {
                return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            var newRow = new Row() {
                    RowIndex = rowIndex
            };
            sheetData.Append(newRow);
            return newRow;
        }

        internal static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            return GetRow(sheetData, rowIndex);
        }

        internal static Row GetRow(WorksheetPart worksheetPart, uint rowIndex)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            return GetRow(sheetData, rowIndex);
        }

        #endregion

        #region copy cell style

        [Obsolete]
        public static void CopyCellStyle(string docName, string sheetName, uint rowIndex, string columnName, string templateSheetName, uint templateRowIndex, string templateColumnName)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                spreadSheet.CopyCellStyle(sheetName, rowIndex, columnName, templateSheetName, templateRowIndex, templateColumnName);
            }
        }

        public static void CopyCellStyle(this SpreadsheetDocument spreadSheet, string sheetName, uint rowIndex, string columnName, string templateSheetName, uint templateRowIndex, string templateColumnName)
        {
            var templateCell = GetCell(spreadSheet, templateSheetName, templateRowIndex, templateColumnName);
            var worksheetPart = GetWorksheetPartByName(spreadSheet, sheetName);
            if (worksheetPart != null)
            {
                var worksheet = worksheetPart.Worksheet;
                var cell = GetCell(worksheet, rowIndex, columnName);
                cell.StyleIndex = templateCell.StyleIndex;
            }
        }

        #endregion

        #region update cell

        [Obsolete]
        public static void UpdateCell(string docName, string sheetName, uint rowIndex, string columnName, string text)
        {
            UpdateCell(docName, sheetName, rowIndex, columnName, CellValues.String, text);
        }

        [Obsolete]
        public static void UpdateCell(string docName, string sheetName, uint rowIndex, string columnName, CellValues dataType, string text)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                spreadSheet.UpdateCell(sheetName, rowIndex, columnName, dataType, text);
            }
        }

        [Obsolete]
        public static void UpdateCell(string docName, string sheetName, uint rowIndex, string columnName, CellValues dataType, string text, string sampleSheetName, uint sampleRowIndex, string sampleColumnIndex)
        {
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                spreadSheet.UpdateCell(sheetName, rowIndex, columnName, dataType, text, sampleSheetName, sampleRowIndex, sampleColumnIndex);
            }
        }

        public static void UpdateCell(this SpreadsheetDocument spreadSheet, string sheetName, uint rowIndex, string columnName, CellValues dataType, string text, string sampleSheetName, uint sampleRowIndex, string sampleColumnIndex)
        {
            uint cellStyleIndex = GetCell(spreadSheet, sampleSheetName, sampleRowIndex, sampleColumnIndex).StyleIndex;
            spreadSheet.UpdateCell(sheetName, rowIndex, columnName, dataType, text, cellStyleIndex);
        }

        internal static Cell GetCell(SpreadsheetDocument spreadsheet, string worksheetName, uint rowIndex, string columnName)
        {
            var worksheetPart = GetWorksheetPartByName(spreadsheet, worksheetName);
            var worksheet = worksheetPart.Worksheet;
            var cell = GetCell(worksheet, rowIndex, columnName);
            return cell;
        }

        public static void UpdateCell(this SpreadsheetDocument spreadSheet, string sheetName, uint rowIndex, string columnName, CellValues dataType, string text)
        {
            var worksheet = GetWorksheetByName(spreadSheet, sheetName);
            if (worksheet != null)
            {
                worksheet.UpdateCell(rowIndex, columnName, dataType, text);
            }
        }

        public static Worksheet GetWorksheetByName(SpreadsheetDocument spreadSheet, string sheetName)
        {
            var worksheetPart = GetWorksheetPartByName(spreadSheet, sheetName);
            Worksheet worksheet = null;
            if (worksheetPart != null)
            {
                worksheet = worksheetPart.Worksheet;
            }
            return worksheet;
        }

        public static void UpdateCell(this SpreadsheetDocument spreadSheet, string sheetName, uint rowIndex, string columnName, CellValues dataType, string text, uint? styleIndex)
        {
            var worksheetPart = GetWorksheetPartByName(spreadSheet, sheetName);
            if (worksheetPart != null)
            {
                var worksheet = worksheetPart.Worksheet;
                var cell = GetCell(worksheet, rowIndex, columnName);
                cell.CellValue = new CellValue(text);
                cell.DataType = new EnumValue<CellValues>(dataType);
                if (styleIndex.HasValue)
                {
                    cell.StyleIndex = styleIndex.Value;
                }
            }
        }

        #endregion

        #region merge cell

        /// <summary>
        ///   Given a document name, a worksheet name, and the names of two adjacent cells, merges the two cells.
        ///   When two cells are merged, only the content from one cell is preserved:
        ///   the upper-left cell for left-to-right languages or the upper-right cell for right-to-left languages.
        /// </summary>
        /// <param name="docName"> </param>
        /// <param name="sheetName"> </param>
        /// <param name="cell1Name"> </param>
        /// <param name="cell2Name"> </param>
        [Obsolete]
        public static void MergeTwoCells(string docName, string sheetName, string cell1Name, string cell2Name)
        {
            // Open the document for editing.
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(docName, true))
            {
                document.MergeTwoCells(sheetName, cell1Name, cell2Name);
            }
        }

        public static void MergeTwoCells(this SpreadsheetDocument document, string sheetName, string cell1Name, string cell2Name)
        {
            var worksheet = GetWorksheet(document, sheetName);
            if (worksheet != null && !string.IsNullOrEmpty(cell1Name) &&
                !string.IsNullOrEmpty(cell2Name))
            {
                CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
                CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

                MergeCells mergeCells;
                if (worksheet.Elements<MergeCells>().Any())
                {
                    mergeCells = worksheet.Elements<MergeCells>().First();
                }
                else
                {
                    mergeCells = new MergeCells();

                    // Insert a MergeCells object into the specified position.
                    if (worksheet.Elements<CustomSheetView>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
                    }
                    else if (worksheet.Elements<DataConsolidate>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
                    }
                    else if (worksheet.Elements<SortState>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
                    }
                    else if (worksheet.Elements<AutoFilter>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
                    }
                    else if (worksheet.Elements<Scenarios>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
                    }
                    else if (worksheet.Elements<ProtectedRanges>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
                    }
                    else if (worksheet.Elements<SheetProtection>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
                    }
                    else if (worksheet.Elements<SheetCalculationProperties>().Any())
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
                    }
                    else
                    {
                        worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
                    }
                }

                // Create the merged cell and append it to the MergeCells collection.
                var mergeCell = new MergeCell() {
                        Reference = new StringValue(cell1Name + ":" + cell2Name)
                };
                mergeCells.Append(mergeCell);
            }
        }

        /// <summary>
        ///   Given a Worksheet and a cell name, verifies that the specified cell exists.
        ///   If it does not exist, creates a new cell.
        /// </summary>
        /// <param name="worksheet"> </param>
        /// <param name="cellName"> </param>
        internal static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            // If the Worksheet does not contain the specified row, create the specified row.
            // Create the specified cell in that row, and insert the row into the Worksheet.
            if (!rows.Any())
            {
                var row = new Row() {
                        RowIndex = new UInt32Value(rowIndex)
                };
                var cell = new Cell() {
                        CellReference = new StringValue(cellName)
                };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
            }
            else
            {
                var row = rows.First();

                var cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                // If the row does not contain the specified cell, create the specified cell.
                if (!cells.Any())
                {
                    var cell = new Cell() {
                            CellReference = new StringValue(cellName)
                    };
                    row.Append(cell);
                }
            }
        }

        /// <summary>
        ///   Given a SpreadsheetDocument and a worksheet name, get the specified worksheet.
        /// </summary>
        /// <param name="document"> </param>
        /// <param name="worksheetName"> </param>
        /// <returns> </returns>
        public static Worksheet GetWorksheet(SpreadsheetDocument document, string worksheetName)
        {
            var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            var worksheetPart = (WorksheetPart) document.WorkbookPart.GetPartById(sheets.First().Id);
            return sheets.Any() ? worksheetPart.Worksheet : null;
        }

        #endregion
    }
}