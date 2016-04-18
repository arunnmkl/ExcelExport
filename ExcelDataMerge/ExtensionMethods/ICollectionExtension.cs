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
        public static void Add(this IList<SetStyle> list, string name, CellStyle headerStyle, CellStyle rowStyle)
        {
            list.Add(new SetStyle()
            {
                HeaderStyle = headerStyle,
                Name = name,
                RowStyle = rowStyle
            });
        }
    }
}
