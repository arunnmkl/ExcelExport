using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataMerge.Enum;

namespace ExcelDataMerge.Model
{
    /// <summary>
    /// Class to encapsulate the excel export model.
    /// </summary>
    public class ExcelExportModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExportModel" /> class.
        /// </summary>
        /// <param name="dataSets">The data sets.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="exportAs">The export as.</param>
        public ExcelExportModel(IList<DataSet> dataSets, string sheetName, ExportType exportAs = ExportType.Merged)
        {
            this.SheetName = sheetName;
            this.ExportAs = exportAs;
            this.DataSets = dataSets;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExportModel" /> class.
        /// </summary>
        /// <param name="dataSet">The data set.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="exportAs">The export as.</param>
        public ExcelExportModel(DataSet dataSet, string sheetName, ExportType exportAs = ExportType.Merged)
        {
            this.SheetName = sheetName;
            this.ExportAs = exportAs;
            this.DataSets = new List<DataSet>()
            {
                dataSet
            };
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExportModel" /> class.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="exportAs">The export as.</param>
        public ExcelExportModel(DataTable dataTable, string sheetName, ExportType exportAs = ExportType.Single)
        {
            this.SheetName = sheetName;
            this.ExportAs = exportAs;
            this.DataSets = new List<DataSet>()
            {
                GetDataSet(dataTable,sheetName)
            };
        }

        /// <summary>
        /// Gets or sets the data sets.
        /// </summary>
        /// <value>
        /// The data sets.
        /// </value>
        public IList<DataSet> DataSets { get; private set; }

        /// <summary>
        /// Gets or sets the name of the sheet.
        /// </summary>
        /// <value>
        /// The name of the sheet.
        /// </value>
        public string SheetName { get; private set; }

        /// <summary>
        /// Gets or sets the export as.
        /// </summary>
        /// <value>
        /// The export as.
        /// </value>
        public ExportType ExportAs { get; private set; }

        /// <summary>
        /// Gets the data set.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>
        /// the defined dataset
        /// </returns>
        private DataSet GetDataSet(DataTable dataTable, string sheetName)
        {
            using (DataSet dataSet = new DataSet(sheetName))
            {
                dataSet.Tables.Add(dataTable);
                return dataSet;
            }
        }
    }
}
