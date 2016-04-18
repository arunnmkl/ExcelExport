using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelDataMerge.Model
{
    /// <summary>
    /// Class to encapsulate the excel export model.
    /// </summary>
    public class ExcelExportModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExportModel"/> class.
        /// </summary>
        public ExcelExportModel()
        {
            this.SetStyles = new List<SetStyle>();
        }

        /// <summary>
        /// Gets or sets the data sets.
        /// </summary>
        /// <value>
        /// The data sets.
        /// </value>
        public IList<DataSet> DataSets { get; set; }

        /// <summary>
        /// Gets or sets the set styles.
        /// </summary>
        /// <value>
        /// The set styles.
        /// </value>
        public IList<SetStyle> SetStyles { get; set; }

        /// <summary>
        /// Gets or sets the name of the sheet.
        /// </summary>
        /// <value>
        /// The name of the sheet.
        /// </value>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets or sets the file path.
        /// </summary>
        /// <value>
        /// The file path.
        /// </value>
        public string FilePath { get; set; }
    }
}
