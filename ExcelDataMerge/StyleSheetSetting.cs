using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDataMerge
{
    /// <summary>
    /// Excel style sheet settings.
    /// </summary>
    public class StyleSheetSetting
    {
        /// <summary>
        /// Gets the fonts.
        /// </summary>
        /// <value>
        /// The fonts.
        /// </value>
        public Fonts Fonts { get; private set; }

        /// <summary>
        /// Gets the fills.
        /// </summary>
        /// <value>
        /// The fills.
        /// </value>
        public Fills Fills { get; private set; }

        /// <summary>
        /// Gets the borders.
        /// </summary>
        /// <value>
        /// The borders.
        /// </value>
        public Borders Borders { get; private set; }

        /// <summary>
        /// Gets the numbering formats.
        /// </summary>
        /// <value>
        /// The numbering formats.
        /// </value>
        public NumberingFormats NumberingFormats { get; private set; }

        /// <summary>
        /// Gets the cell formats.
        /// </summary>
        /// <value>
        /// The cell formats.
        /// </value>
        public CellFormats CellFormats { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="StyleSheetSetting" /> class.
        /// </summary>
        public StyleSheetSetting()
        {
            this.Fonts = this.ReadFonts();
            this.Fills = this.ReadFills();
            this.Borders = this.ReadBorders();
            this.NumberingFormats = this.ReadNumberingFormats();
            this.CellFormats = this.ReadCellFormats();
        }

        /// <summary>
        /// Reads the fonts.
        /// </summary>
        /// <returns>
        /// the fonts
        /// </returns>
        private Fonts ReadFonts()
        {
            return new Fonts(
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
            );
        }

        /// <summary>
        /// Reads the fills.
        /// </summary>
        /// <returns>
        /// the fills
        /// </returns>
        private Fills ReadFills()
        {
            return new Fills(
                new Fill(                                                           // Index 0 - The default fill.
                    new PatternFill() { PatternType = PatternValues.None }),
                new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                    new PatternFill() { PatternType = PatternValues.Gray125 }),
                new Fill(                                                           // Index 2 - The yellow fill.
                    new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }) { PatternType = PatternValues.Solid }),
                new Fill(                                                           // Index 3 - The lite blue fill.
                    new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "8DB3E2" } }) { PatternType = PatternValues.Solid }),
                new Fill(                                                           // Index 4 - The navy fill.
                    new PatternFill(new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FBD4B4" } }) { PatternType = PatternValues.Solid })
            );
        }

        /// <summary>
        /// Reads the borders.
        /// </summary>
        /// <returns>
        /// the borders
        /// </returns>
        private Borders ReadBorders()
        {
            return new Borders(
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
            );
        }

        private NumberingFormats ReadNumberingFormats()
        {
            uint iExcelIndex = 164;
            return new NumberingFormats(
              new NumberingFormat() { NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++), FormatCode = StringValue.FromString("dd/mm/yyyy hh:mm:ss") }, // Date time -- 164
              new NumberingFormat() { NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++), FormatCode = StringValue.FromString("#,##0.0000") },  // 4 decimal -- 165
              new NumberingFormat() { NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++), FormatCode = StringValue.FromString("#,##0.00") },    // 2 decimal -- 166
              new NumberingFormat() { NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++), FormatCode = StringValue.FromString("@") },   // ForcedText -- 167
              new NumberingFormat() { NumberFormatId = UInt32Value.FromUInt32(iExcelIndex++), FormatCode = StringValue.FromString("\"" + System.Globalization.CultureInfo.CurrentUICulture.NumberFormat.CurrencySymbol + "\"\\ " + "#,##0.00") }   // Currency -- 168
            );
        }

        /// <summary>
        /// Reads the cell formats.
        /// </summary>
        /// <returns>
        /// the cell formats
        /// </returns>
        private CellFormats ReadCellFormats()
        {
            return new CellFormats(
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 },                         // Index 0 - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                    new CellFormat() { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 1 - Bold 
                    new CellFormat() { FontId = 2, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 2 - Italic
                    new CellFormat() { FontId = 3, FillId = 0, BorderId = 0, ApplyFont = true },       // Index 3 - Times Roman
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 2, BorderId = 1, ApplyFill = true, ApplyAlignment = true },       // Index 4 - Yellow Fill
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 0, ApplyAlignment = true },     // Index 5 - Alignment
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true },     // Index 6 - Border
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 3, BorderId = 1, ApplyFill = true, ApplyAlignment = true },       // Index 7 - lite blue Fill
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 4, BorderId = 1, ApplyFill = true, ApplyAlignment = true },       // Index 8 - navy Fill
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }) { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, ApplyAlignment = true },     // Index 9 - AlignmentWithBorder
                    new CellFormat() { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, NumberFormatId = 167, ApplyNumberFormat = true }     // Index 10 - Border Currency
            );
        }
    }
}
