using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDataMerge.Enum;

namespace ExcelDataMerge.Model
{
    /// <summary>
    /// To encapsulate the set styles.
    /// </summary>
    public class SetStyle
    {
        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the header style.
        /// </summary>
        /// <value>
        /// The header style.
        /// </value>
        public CellStyle HeaderStyle { get; set; }

        /// <summary>
        /// Gets or sets the row style.
        /// </summary>
        /// <value>
        /// The row style.
        /// </value>
        public CellStyle RowStyle { get; set; }

        /// <summary>
        /// Gets or sets the index of the column.
        /// </summary>
        /// <value>
        /// The index of the column.
        /// </value>
        public int? ColumnIndex { get; set; }
    }
}
