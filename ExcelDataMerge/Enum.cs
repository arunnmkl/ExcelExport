using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelDataMerge.Enum
{
    /// <summary>
    /// Enums for Cell style.
    /// </summary>
    public enum CellStyle
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
        Border = 6,
        /// <summary>
        /// The lite blue fill
        /// </summary>
        LiteBlueFill = 7,
        /// <summary>
        /// The navy fill
        /// </summary>
        NavyFill = 8,
        /// <summary>
        /// The alignment with border
        /// </summary>
        AlignmentWithBorder = 9
    }

    /// <summary>
    ///  Enums for Cell type.
    /// </summary>
    public enum CellType
    {
        /// <summary>
        /// The header
        /// </summary>
        Header,
        /// <summary>
        /// The row
        /// </summary>
        Row
    }
}
