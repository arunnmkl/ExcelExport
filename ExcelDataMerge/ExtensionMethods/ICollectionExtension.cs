using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDataMerge.Enum;
using ExcelDataMerge.Model;

namespace ExcelDataMerge.ExtensionMethods
{
    public static class ICollectionExtension
    {
        /// <summary>
        /// Adds the specified name.
        /// </summary>
        /// <param name="list">The list.</param>
        /// <param name="name">The name.</param>
        /// <param name="headerStyle">The header style.</param>
        /// <param name="rowStyle">The row style.</param>
        /// <param name="columnIndex">Index of the column.</param>
        public static void Add(this IList<SetStyle> list, string name, CellStyle headerStyle, CellStyle rowStyle, int? columnIndex = null)
        {
            if (!list.Any(l => l.Name == name && (columnIndex.HasValue && l.ColumnIndex.HasValue ? l.ColumnIndex.Value == columnIndex.Value : columnIndex == l.ColumnIndex)))
            {
                list.Add(new SetStyle()
                {
                    HeaderStyle = headerStyle,
                    Name = name,
                    RowStyle = rowStyle,
                    ColumnIndex = columnIndex
                });

            }
        }
    }
}
