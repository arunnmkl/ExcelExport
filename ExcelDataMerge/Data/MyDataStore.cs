using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataMerge.Data
{
    /// <summary>
    /// My data store
    /// </summary>
    public class MyDataStore
    {
        /// <summary>
        /// Gets the multiple sets.
        /// </summary>
        /// <param name="setCount">The set count.</param>
        /// <param name="tableCount">The table count.</param>
        /// <returns>
        /// dataset collection
        /// </returns>
        public static IList<DataSet> GetMultipleSets(int setCount = 1, int tableCount = 3)
        {
            IList<DataSet> dsList = new List<DataSet>();

            for (int index = 0; index < setCount; index++)
            {
                DataSet ds = InitDataSet(index + 1, tableCount);

                dsList.Add(ds);
            }

            return dsList;
        }

        /// <summary>
        /// Gets the data set.
        /// </summary>
        /// <param name="tableCount">The table count.</param>
        /// <returns>
        /// data set
        /// </returns>
        public static DataSet GetDataSet(int tableCount = 3)
        {
            return InitDataSet(tableCount: tableCount);
        }

        /// <summary>
        /// Gets the table data.
        /// </summary>
        /// <param name="tableName">Name of the table.</param>
        /// <param name="columnCount">The column count.</param>
        /// <param name="recordCount">The record count.</param>
        /// <returns>
        /// data table
        /// </returns>
        public static DataTable GetTableData(string tableName, int columnCount = 4, int recordCount = 10)
        {
            if (columnCount < 2)
            {
                columnCount = 2;
            }
            var dataTable = new DataTable(tableName);

            DataColumn dc1 = new DataColumn("Column 1", DbType.Int32.GetType());
            dc1.AutoIncrement = true;
            dc1.AutoIncrementSeed = new Random().Next(98);
            dataTable.Columns.Add(dc1);

            DataColumn dc2 = new DataColumn("Column 2", typeof(decimal));
            dataTable.Columns.Add(dc2);

            for (int i = 3; i <= columnCount; i++)
            {
                dataTable.Columns.Add(InitColumn(i));
            }

            for (int i = 0; i < recordCount; i++)
            {
                var number = new Random().Next(6 * (i + 2));
                var dr = dataTable.NewRow();
                dr["Column 2"] = number;

                for (int cIndex = 3; cIndex <= columnCount; cIndex++)
                {
                    dr[string.Format("Column {0}", cIndex)] = cIndex % 2 == 0 ? RandomString(number % 7 == 0 ? number / 7 : number % 7) : RandomString(number % 4 == 0 ? number / 4 : number % 4);
                }

                dataTable.Rows.Add(dr);
            }

            return dataTable;
        }

        /// <summary>
        /// Initializes the data set.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="tableCount">The table count.</param>
        /// <returns>
        /// data set
        /// </returns>
        private static DataSet InitDataSet(int index = 1, int tableCount = 3)
        {
            DataSet ds = new DataSet(string.Concat("Set ", index));

            if (tableCount == 3 || tableCount == 0)
            {
                ds.Tables.Add(GetTableData("T One", recordCount: 2));

                ds.Tables.Add(GetTableData("T Two", columnCount: 5, recordCount: 4));

                ds.Tables.Add(GetTableData("T Three", columnCount: 7, recordCount: 6));
            }
            else
            {
                for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
                {
                    ds.Tables.Add(GetTableData(string.Concat("Table ", tableIndex), recordCount: tableIndex * tableCount * 2));
                }
            }

            return ds;
        }

        /// <summary>
        /// Initializes the column.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns>
        /// the initiated columns
        /// </returns>
        private static DataColumn InitColumn(int index)
        {
            return new DataColumn(string.Format("Column {0}", index), typeof(string));
        }

        /// <summary>
        /// Random string.
        /// </summary>
        /// <param name="length">The length.</param>
        /// <returns>
        /// the random string
        /// </returns>
        private static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
            var random = new Random();
            return new string(Enumerable.Repeat(chars, length / 2).Select(s => s[random.Next(s.Length / 2)]).ToArray());
        }
    }

    /// <summary>
    /// Helper class of my data store
    /// </summary>
    public class MyDataStoreHelper
    {
        /// <summary>
        /// Gets the header name list.
        /// </summary>
        /// <param name="dataSet">The data set.</param>
        /// <returns>
        /// combination of table name and its columns as list
        /// </returns>
        public static IDictionary<string, IList<string>> GetHeaderNameList(DataSet dataSet)
        {
            IDictionary<string, IList<string>> headerNameList = new Dictionary<string, IList<string>>();
            if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0)
            {
                return headerNameList;
            }

            for (int index = 0; index < dataSet.Tables.Count; index++)
            {
                IList<string> columnNames = new List<string>();
                foreach (DataColumn column in dataSet.Tables[index].Columns)
                {
                    columnNames.Add(column.ColumnName);
                }

                headerNameList.Add(dataSet.Tables[index].TableName, columnNames);
            }

            return headerNameList;
        }

        /// <summary>
        /// Converts to row data list.
        /// </summary>
        /// <param name="dataSet">The data set.</param>
        /// <returns>
        /// combination of table name and its row values
        /// </returns>
        public static IDictionary<string, IList<object[]>> ConvertToRowDataList(DataSet dataSet)
        {
            IDictionary<string, IList<object[]>> rowDataList = new Dictionary<string, IList<object[]>>();
            if (dataSet == null || dataSet.Tables == null || dataSet.Tables.Count == 0)
            {
                return rowDataList;
            }

            for (int index = 0; index < dataSet.Tables.Count; index++)
            {
                IList<object[]> rowItems = new List<object[]>();
                foreach (DataRow row in dataSet.Tables[index].Rows)
                {
                    rowItems.Add(row.ItemArray);
                }

                rowDataList.Add(dataSet.Tables[index].TableName, rowItems);
            }

            return rowDataList;
        }

        /// <summary>
        /// Gets the header name list.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <returns>column name as in string array</returns>
        public static IList<string> GetHeaderNameList(DataTable dataTable)
        {
            IList<string> columnNameList = new List<string>();
            if (dataTable == null || dataTable.Columns == null || dataTable.Columns.Count == 0)
            {
                return columnNameList;
            }

            for (int index = 0; index < dataTable.Columns.Count; index++)
            {
                columnNameList.Add(dataTable.Columns[index].ColumnName);
            }

            return columnNameList;
        }

        /// <summary>
        /// Converts to row data list.
        /// </summary>
        /// <param name="dataTable">The data table.</param>
        /// <returns>row values as in object array</returns>
        public static IList<object[]> ConvertToRowDataList(DataTable dataTable)
        {
            IList<object[]> rowItemList = new List<object[]>();
            if (dataTable == null || dataTable.Rows == null || dataTable.Rows.Count == 0)
            {
                return rowItemList;
            }

            for (int index = 0; index < dataTable.Rows.Count; index++)
            {
                rowItemList.Add(dataTable.Rows[index].ItemArray);
            }

            return rowItemList;
        }
    }
}
