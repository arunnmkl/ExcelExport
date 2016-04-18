using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataMerge.Data;

namespace ExcelDataMerge
{
    /// <summary>
    /// Sample development for single excel sheet with multiple data collection.
    /// </summary>
    public class DynamicDataIntoExcel
    {
        /// <summary>
        /// Creates the spreadsheet workbook.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="dataSets">The data sets.</param>
        public static void CreateSpreadsheetWorkbook(string filePath, string sheetName, IList<DataSet> dataSets)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                // Append a new worksheet and associate it with the workbook.
                sheets.Append(new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                });

                IList<MergeCell> mergeCellItem = new List<MergeCell>();

                // Add header
                UInt32 rowIndex = 0;
                int cellIndex = 0;
                int initialCellIndex = cellIndex;
                Row row = new Row { RowIndex = ++rowIndex };
                sheetData.AppendChild(row);

                foreach (DataSet data in dataSets)
                {
                    CreateSheetData(mergeCellItem, sheetData, out rowIndex, ref cellIndex, out initialCellIndex, ref row, data);
                }

                SetMergeCell(mergeCellItem, worksheetPart);

                ApplyStyles(spreadsheetDocument);

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
            }
        }

        /// <summary>
        /// Creates the spreadsheet workbook.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="data">The data.</param>
        public static void CreateSpreadsheetWorkbook(string filePath, string sheetName, DataSet data)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                // Append a new worksheet and associate it with the workbook.
                sheets.Append(new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = sheetName
                });

                IList<MergeCell> mergeCellItem = new List<MergeCell>();

                // Add header
                UInt32 rowIndex = 0;
                int cellIndex = 0;
                int initialCellIndex = cellIndex;
                Row row = new Row { RowIndex = ++rowIndex };
                sheetData.AppendChild(row);

                CreateSheetData(mergeCellItem, sheetData, out rowIndex, ref cellIndex, out initialCellIndex, ref row, data);
                SetMergeCell(mergeCellItem, worksheetPart);
                ApplyStyles(spreadsheetDocument);

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();
            }
        }

        /// <summary>
        /// Creates the sheet data.
        /// </summary>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="cellIndex">Index of the cell.</param>
        /// <param name="initialCellIndex">Initial index of the cell.</param>
        /// <param name="row">The row.</param>
        /// <param name="data">The data.</param>
        private static void CreateSheetData(IList<MergeCell> mergeCellItem, SheetData sheetData, out uint rowIndex, ref int cellIndex, out int initialCellIndex, ref Row row, DataSet data)
        {
            rowIndex = 0;
            ++rowIndex;
            initialCellIndex = cellIndex;
            IDictionary<string, IList<string>> headerNameList = MyDataStoreHelper.GetHeaderNameList(data);
            IDictionary<string, IList<object[]>> rowDataList = MyDataStoreHelper.ConvertToRowDataList(data);

            ApplyHeader(mergeCellItem, headerNameList, data.DataSetName, ref rowIndex, row, ref cellIndex, initialCellIndex);
            AddRows(sheetData, ref rowIndex, ref cellIndex, ref row, rowDataList, initialCellIndex);
        }

        /// <summary>
        /// Applies the header.
        /// </summary>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="headerNameList">The header name list.</param>
        /// <param name="dataSetName">Name of the data set.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="row">The row.</param>
        /// <param name="cellIndex">Index of the cell.</param>
        private static void ApplyHeader(IList<MergeCell> mergeCellItem, IDictionary<string, IList<string>> headerNameList, string dataSetName, ref uint rowIndex, Row row, ref int cellIndex, int initialCellIndex)
        {
            IList<string> indexer = new List<string>();

            if (string.IsNullOrEmpty(dataSetName) == false)
            {
                goto NewHeader;
            }
            else
            {
                goto NewHeader;
            }

            NewHeader:
            foreach (string master in headerNameList.Keys)
            {
                var mergeCellHeader = headerNameList[master];
                string clFirst = ColumnLetter(cellIndex++);
                string clLast = clFirst;
                foreach (var item in mergeCellHeader)
                {
                    row.AppendChild(CreateTextCell(clLast, rowIndex, dataSetName ?? master ?? string.Empty, (UInt32)CellStyleEnum.Alignment));
                    // check for if not the last iteration of the loop
                    if (mergeCellHeader.IndexOf(item) != mergeCellHeader.Count - 1)
                    {
                        clLast = ColumnLetter(cellIndex++);
                    }
                }

                if (string.IsNullOrEmpty(dataSetName) == true)
                {
                    // Create the merged cell and append it to the MergeCells collection.
                    mergeCellItem.Add(new MergeCell()
                    {
                        Reference = new StringValue(string.Concat(clFirst, rowIndex) + ":" + string.Concat(clLast, rowIndex))
                    });
                }

                indexer.Add(clFirst);
                indexer.Add(clLast);
            }

            if (string.IsNullOrEmpty(dataSetName) == false)
            {
                dataSetName = null;

                // Create the merged cell and append it to the MergeCells collection.
                mergeCellItem.Add(new MergeCell()
                {
                    Reference = new StringValue(string.Concat(indexer.First(), rowIndex) + ":" + string.Concat(indexer.Last(), rowIndex))
                });

                rowIndex++;
                cellIndex = initialCellIndex;

                goto NewHeader;
            }

            rowIndex++;
            cellIndex = initialCellIndex;
            foreach (var headers in headerNameList)
            {
                foreach (string header in headers.Value)
                {
                    row.AppendChild(CreateTextCell(ColumnLetter(cellIndex++), rowIndex, header ?? string.Empty, (UInt32)CellStyleEnum.Alignment));
                }
            }
        }

        /// <summary>
        /// Adds the rows.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="cellIndex">Index of the cell.</param>
        /// <param name="row">The row.</param>
        /// <param name="rowDataList">The row data list.</param>
        private static void AddRows(SheetData sheetData, ref uint rowIndex, ref int cellIndex, ref Row row, IDictionary<string, IList<object[]>> rowDataList, int initialCellIndex)
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
                    prevCellIndex += item.Value.FirstOrDefault().Count();
                }
                else
                {
                    prevCellIndex += 0;
                }
            }
        }

        /// <summary>
        /// Sets the merge cell.
        /// </summary>
        /// <param name="mergeCellItem">The merge cell item.</param>
        /// <param name="worksheetPart">The worksheet part.</param>
        private static void SetMergeCell(IList<MergeCell> mergeCellItem, WorksheetPart worksheetPart)
        {
            if (worksheetPart.Worksheet.Elements<MergeCells>().Count() == 0)
            {
                MergeCells mergeCells = new MergeCells();

                // Insert a MergeCells object into the specified position.
                if (worksheetPart.Worksheet.Elements<CustomSheetView>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<CustomSheetView>().First());
                }
                else if (worksheetPart.Worksheet.Elements<DataConsolidate>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<DataConsolidate>().First());
                }
                else if (worksheetPart.Worksheet.Elements<SortState>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SortState>().First());
                }
                else if (worksheetPart.Worksheet.Elements<AutoFilter>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<AutoFilter>().First());
                }
                else if (worksheetPart.Worksheet.Elements<Scenarios>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<Scenarios>().First());
                }
                else if (worksheetPart.Worksheet.Elements<ProtectedRanges>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<ProtectedRanges>().First());
                }
                else if (worksheetPart.Worksheet.Elements<SheetProtection>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetProtection>().First());
                }
                else if (worksheetPart.Worksheet.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                }

                mergeCells.Append(mergeCellItem.ToArray());
            }
        }

        /// <summary>
        /// Applies the styles.
        /// </summary>
        /// <param name="spreadsheetDocument">The spreadsheet document.</param>
        private static void ApplyStyles(SpreadsheetDocument spreadsheetDocument)
        {
            WorkbookStylesPart stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = GenerateStyleSheet();
            stylesPart.Stylesheet.Save();
        }

        /// <summary>
        /// Enums for Cell style.
        /// </summary>
        private enum CellStyleEnum
        {
            /// <summary>
            /// The default
            /// </summary>
            Default = 0,
            /// <summary>
            /// The bold
            /// </summary>
            Bold = 1,
            /// <summary>
            /// The italic
            /// </summary>
            Italic = 2,
            /// <summary>
            /// The times roman
            /// </summary>
            TimesRoman = 3,
            /// <summary>
            /// The yellow fill
            /// </summary>
            YellowFill = 4,
            /// <summary>
            /// The alignment
            /// </summary>
            Alignment = 5,
            /// <summary>
            /// The border
            /// </summary>
            Border = 6
        }

        /// <summary>
        /// Generates the style sheet.
        /// </summary>
        /// <returns>style sheet</returns>
        private static Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 1 - The bold font.
                        new Bold(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Italic font.
                        new Italic(),
                        new FontSize() { Val = 11 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Calibri" }),
                    new Font(                                                               // Index 2 - The Times Roman font. with 16 size
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })
                ),
                new Fills(
                    new Fill(                                                           // Index 0 - The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(                                                           // Index 2 - The yellow fill.
                        new PatternFill(
                            new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                        )
                        { PatternType = PatternValues.Solid })
                ),
                new Borders(
                    new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),                                                         // Index 0 - The default border.
                    new Border(
                        new LeftBorder(new Color() { Auto = true })
                        {
                            Style = BorderStyleValues.Thin
                        }
                        , new RightBorder(new Color() { Auto = true })
                        {
                            Style = BorderStyleValues.Thin
                        }
                        , new TopBorder(new Color() { Auto = true })
                        {
                            Style = BorderStyleValues.Thin
                        }
                        , new BottomBorder(new Color() { Auto = true })
                        {
                            Style = BorderStyleValues.Thin
                        }
                        , new DiagonalBorder()) // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                ),
                new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                         // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 - Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 - Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 - Times Roman
                    new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true },       // Index 4 - Yellow Fill
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },     // Index 5 - Alignment
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true }      // Index 6 - Border
                )
            );
        }

        /// <summary>
        /// Columns the letter.
        /// </summary>
        /// <param name="columnXAxis">The column x axis.</param>
        /// <returns>position identity</returns>
        private static string ColumnLetter(int columnXAxis)
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
        /// <returns>initiated cell</returns>
        private static Cell CreateTextCell(string header, UInt32 index, string text, UInt32 styleIndex = (UInt32)CellStyleEnum.Default)
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
        /// <returns>initiated cell</returns>
        private static Cell CreateTextCell(string header, UInt32 index, object text, UInt32 styleIndex = (UInt32)CellStyleEnum.Default)
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
    }
}
