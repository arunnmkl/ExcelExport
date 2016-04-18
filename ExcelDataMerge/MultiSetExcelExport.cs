using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataMerge.Data;
using ExcelDataMerge.Model;

namespace ExcelDataMerge
{
    /// <summary>
    /// 
    /// </summary>
    public class MultiSetExcelExport
    {
        /// <summary>
        /// The row index
        /// </summary>
        private static uint rowIndex;

        /// <summary>
        /// The cell index
        /// </summary>
        private static uint cellIndex;

        /// <summary>
        /// The row
        /// </summary>
        private static Row row;

        /// <summary>
        /// The initial cell index
        /// </summary>
        private static uint initialCellIndex;

        /// <summary>
        /// The set styles
        /// </summary>
        private static IList<SetStyle> setStyles = new List<SetStyle>();

        /// <summary>
        /// Creates the excel document.
        /// </summary>
        /// <param name="model">The model.</param>
        /// <returns>
        /// export state
        /// </returns>
        /// <exception cref="System.ArgumentNullException">model</exception>
        internal static bool CreateExcelDocument(ExcelExportModel model)
        {
            if (model == null)
            {
                throw new ArgumentNullException(nameof(model));
            }

            setStyles = model.SetStyles;
            return CreateExcelDocument(model.FilePath, model.SheetName, model.DataSets);
        }

        /// <summary>
        /// Creates the excel document.
        /// </summary>
        /// <param name="excelFilePath">The excel file path.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="dataSets">The data sets.</param>
        /// <returns>export state</returns>
        internal static bool CreateExcelDocument(string excelFilePath, string sheetName, IList<DataSet> dataSets)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook))
                {
                    WriteExcelFile(dataSets, spreadsheetDocument, sheetName);
                    return true;
                }
            }
            catch (Exception ex)
            {
                // ::TODO to add logs.

                return false;
            }
        }

        /// <summary>
        /// Writes the excel file.
        /// </summary>
        /// <param name="dataSets">The data sets.</param>
        /// <param name="spreadsheetDocument">The spreadsheet document.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        private static void WriteExcelFile(IList<DataSet> dataSets, SpreadsheetDocument spreadsheetDocument, string sheetName)
        {
            //  Create the Excel file contents.  This function is used when creating an Excel file either writing 
            //  to a file, or writing to a MemoryStream.
            spreadsheetDocument.AddWorkbookPart();
            spreadsheetDocument.WorkbookPart.Workbook = new Workbook();

            // the following line of code (which prevents crashes in Excel 2010)
            spreadsheetDocument.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

            WorkbookStylesPart stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet();

            uint worksheetNumber = 1;
            CreateNewSheet(dataSets, spreadsheetDocument, sheetName, worksheetNumber);

            spreadsheetDocument.WorkbookPart.Workbook.Save();
        }

        /// <summary>
        /// Creates the new sheet.
        /// </summary>
        /// <param name="dataSets">The data sets.</param>
        /// <param name="spreadsheetDocument">The spreadsheet document.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="worksheetNumber">The worksheet number.</param>
        private static void CreateNewSheet(IList<DataSet> dataSets, SpreadsheetDocument spreadsheetDocument, string sheetName, uint worksheetNumber)
        {
            // New sheet creation and appending the data into it.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet();

            // create sheet data
            newWorksheetPart.Worksheet.AppendChild(new SheetData());

            IList<MergeCell> mergeCellItem = new List<MergeCell>();

            foreach (DataSet data in dataSets)
            {
                CreateSheetData(newWorksheetPart, mergeCellItem, data);
            }

            SetMergeCell(mergeCellItem, newWorksheetPart);

            newWorksheetPart.Worksheet.Save();

            // create the worksheet to workbook relation
            if (worksheetNumber == 1)
            {
                spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
            }

            spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = worksheetNumber,
                Name = sheetName
            });
        }

        /// <summary>
        /// Creates the sheet data.
        /// </summary>
        /// <param name="newWorksheetPart">The new worksheet part.</param>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="data">The data.</param>
        private static void CreateSheetData(WorksheetPart newWorksheetPart, IList<MergeCell> mergeCellItem, DataSet data)
        {
            var worksheet = newWorksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();

            rowIndex = 0;
            row = new Row { RowIndex = ++rowIndex };
            sheetData.AppendChild(row);
            initialCellIndex = cellIndex;

            IDictionary<string, IList<string>> headerNameList = MyDataStoreHelper.GetHeaderNameList(data);
            IDictionary<string, IList<object[]>> rowDataList = MyDataStoreHelper.ConvertToRowDataList(data);

            ApplyHeader(mergeCellItem, headerNameList, data.DataSetName);
            AddRows(sheetData, rowDataList);
        }

        /// <summary>
        /// Sets the merge cell.
        /// </summary>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="newWorksheetPart">The new worksheet part.</param>
        private static void SetMergeCell(IList<MergeCell> mergeCellItem, WorksheetPart newWorksheetPart)
        {
            if (newWorksheetPart.Worksheet.Elements<MergeCells>().Count() == 0)
            {
                MergeCells mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (newWorksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<CustomSheetView>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<DataConsolidate>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<SortState>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<SortState>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<AutoFilter>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<Scenarios>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<Scenarios>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<ProtectedRanges>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<SheetProtection>().First());
                }
                else if (newWorksheetPart.Worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    newWorksheetPart.Worksheet.InsertAfter(mergeCells, newWorksheetPart.Worksheet.Elements<SheetData>().First());
                }

                mergeCells.Append(mergeCellItem.ToArray());
            }
        }

        /// <summary>
        /// Applies the header.
        /// </summary>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="headerNameList">The header name list.</param>
        /// <param name="dataSetName">Name of the data set.</param>
        private static void ApplyHeader(IList<MergeCell> mergeCellItem, IDictionary<string, IList<string>> headerNameList, string dataSetName)
        {
            IList<string> indexerPrefix = new List<string>();
            IList<string> fullIndexer = new List<string>();
            uint? cellStyle = null;

            if (string.IsNullOrEmpty(dataSetName) == false)
            {
                goto NewHeader;
            }

            NewHeader:
            foreach (string master in headerNameList.Keys)
            {
                var mergeCellHeader = headerNameList[master];
                string clFirst = ColumnLetter(cellIndex++);
                string clLast = clFirst;
                if (headerNameList.Count == 1 && dataSetName != null)
                {
                    dataSetName = master;
                }

                if (cellStyle.HasValue == false)
                {
                    cellStyle = GetCellStyle(dataSetName, Enum.CellType.Header);
                }

                foreach (var item in mergeCellHeader)
                {
                    row.AppendChild(CreateTextCell(clLast, rowIndex, dataSetName ?? master ?? string.Empty, cellStyle.Value));
                    // check for if not the last iteration of the loop
                    if (mergeCellHeader.IndexOf(item) != mergeCellHeader.Count - 1)
                    {
                        clLast = ColumnLetter(cellIndex++);
                    }
                }

                if (string.IsNullOrEmpty(dataSetName) && headerNameList.Count != 1)
                {
                    // Create the merged cell and append it to the MergeCells collection.
                    mergeCellItem.Add(new MergeCell()
                    {
                        Reference = new StringValue(string.Concat(clFirst, rowIndex) + ":" + string.Concat(clLast, rowIndex))
                    });
                }

                indexerPrefix.Add(clFirst);
                indexerPrefix.Add(clLast);
                fullIndexer.Add(string.Concat(clFirst, rowIndex));
                fullIndexer.Add(string.Concat(clLast, rowIndex));
            }

            if (string.IsNullOrEmpty(dataSetName) == false)
            {
                dataSetName = null;
                if (headerNameList.Count != 1)
                {
                    // Create the merged cell and append it to the MergeCells collection.
                    mergeCellItem.Add(new MergeCell()
                    {
                        Reference = new StringValue(string.Concat(indexerPrefix.First(), rowIndex) + ":" + string.Concat(indexerPrefix.Last(), rowIndex))
                    });
                }

                rowIndex++;
                cellIndex = initialCellIndex;

                goto NewHeader;
            }

            if (headerNameList.Count == 1)
            {
                // Create the merged cell and append it to the MergeCells collection.
                mergeCellItem.Add(new MergeCell()
                {
                    Reference = new StringValue(fullIndexer.First() + ":" + fullIndexer.Last())
                });
            }

            rowIndex++;
            cellIndex = initialCellIndex;
            foreach (var headers in headerNameList)
            {
                foreach (string header in headers.Value)
                {
                    row.AppendChild(CreateTextCell(ColumnLetter(cellIndex++), rowIndex, header ?? string.Empty, cellStyle.Value));
                }
            }
        }

        /// <summary>
        /// Adds the rows.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowDataList">The row data list.</param>
        private static void AddRows(SheetData sheetData, IDictionary<string, IList<object[]>> rowDataList)
        {
            var prevRowIndex = rowIndex;
            var prevCellIndex = initialCellIndex;
            foreach (var item in rowDataList)
            {
                rowIndex = prevRowIndex;
                // Add sheet data
                foreach (var rowData in item.Value)
                {
                    cellIndex = prevCellIndex;
                    row = new Row { RowIndex = ++rowIndex };
                    sheetData.AppendChild(row);
                    foreach (var cellData in rowData)
                    {
                        var cell = CreateTextCell(ColumnLetter(cellIndex++), rowIndex, cellData ?? string.Empty);
                        row.AppendChild(cell);
                    }
                }

                if (item.Value != null)
                {
                    prevCellIndex += (uint)item.Value.FirstOrDefault().Count();
                }
                else
                {
                    prevCellIndex += 0;
                }
            }
        }

        /// <summary>
        /// Columns the letter.
        /// </summary>
        /// <param name="columnXAxis">The column x axis.</param>
        /// <returns>
        /// position identity
        /// </returns>
        private static string ColumnLetter(uint columnXAxis)
        {
            var intFirstLetter = ((columnXAxis) / 676) + 64;
            var intSecondLetter = ((columnXAxis % 676) / 26) + 64;
            var intThirdLetter = (columnXAxis % 26) + 65;

            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }

        /// <summary>
        /// Creates the text cell.
        /// </summary>
        /// <param name="header">The header.</param>
        /// <param name="index">The index.</param>
        /// <param name="text">The text.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>
        /// initiated cell
        /// </returns>
        private static Cell CreateTextCell(string header, uint index, string text, uint styleIndex = (uint)Enum.CellStyle.Default)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = header + index,
                StyleIndex = styleIndex
            };

            var ilString = new InlineString();
            ilString.AppendChild(new Text { Text = text });
            cell.AppendChild(ilString);
            return cell;
        }

        /// <summary>
        /// Creates the text cell.
        /// </summary>
        /// <param name="header">The header.</param>
        /// <param name="index">The index.</param>
        /// <param name="text">The text.</param>
        /// <param name="styleIndex">Index of the style.</param>
        /// <returns>
        /// initiated cell
        /// </returns>
        private static Cell CreateTextCell(string header, uint index, object text, uint styleIndex = (uint)Enum.CellStyle.Border)
        {
            Cell cell = new Cell
            {
                CellReference = header + index,
                StyleIndex = styleIndex
            };

            CellValue cellvalue = new CellValue()
            {
                Text = Convert.ToString(text)
            };

            if (text is Int16
                || text is Int32
                || text is Int64
                || text is Double
                || text is Decimal)
            {
                cell.DataType = CellValues.Number;
            }
            else if (text is bool)
            {
                cell.DataType = CellValues.Boolean;
            }
            else if (text is DateTime)
            {
                cell.DataType = CellValues.Date;
            }
            else
            {
                cell.DataType = CellValues.String;
            }

            cell.Append(cellvalue);
            return cell;
        }

        /// <summary>
        /// Generates the style sheet.
        /// </summary>
        /// <returns>
        /// style sheet
        /// </returns>
        private static Stylesheet GenerateStyleSheet()
        {
            StyleSheetSetting eStyleSheet = new StyleSheetSetting();
            return new Stylesheet(eStyleSheet.Fonts, eStyleSheet.Fills, eStyleSheet.Borders, eStyleSheet.CellFormats);
        }

        /// <summary>
        /// Gets the cell style.
        /// </summary>
        /// <param name="dataSetName">Name of the data set.</param>
        /// <param name="cellType">Type of the cell.</param>
        /// <returns>cell style</returns>
        private static uint GetCellStyle(string dataSetName, Enum.CellType cellType)
        {
            if (setStyles != null)
            {
                var setStyle = setStyles.FirstOrDefault(ss => ss.Name == dataSetName);
                if (setStyle != null)
                {
                    switch (cellType)
                    {
                        case Enum.CellType.Header:
                            return (uint)setStyle.HeaderStyle;
                        case Enum.CellType.Row:
                            return (uint)setStyle.RowStyle;
                        default:
                            return (uint)Enum.CellStyle.Default;
                    }
                }
            }

            switch (cellType)
            {
                case Enum.CellType.Header:
                    return (uint)Enum.CellStyle.AlignmentWithBorder;
                default:
                    return (uint)Enum.CellStyle.Border;
            }
        }
    }
}
