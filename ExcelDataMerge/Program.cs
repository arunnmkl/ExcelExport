using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataMerge.Data;
using ExcelDataMerge.Model;
using ExcelDataMerge.ExtensionMethods;
using ExcelDataMerge.Enum;

namespace ExcelDataMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            IList<DataSet> dataSets = MyDataStore.GetMultipleSets(3, 1);
            dataSets.Add(MyDataStore.GetDataSet(2));

            IList<SetStyle> setStyles = new List<SetStyle>();
            setStyles.Add("Set 1", CellStyle.YellowFill, CellStyle.Default);
            setStyles.Add("Set 2", CellStyle.NavyFill, CellStyle.Default);
            setStyles.Add("Table 1", CellStyle.LiteBlueFill, CellStyle.Default);

            ExcelExportModel model = new ExcelExportModel
            {
                DataSets = dataSets,
                FilePath = @"C:\AKK\Excel\Multiple_New.xls",
                SheetName = @"Multi Set Data",
                SetStyles = setStyles
            };

            MultiSetExcelExport.CreateExcelDocument(model);
            //MultiSetExcelExport.CreateExcelDocument(@"C:\AKK\Excel\Multiple_New.xls", "Multi Set Data", dataSets);
        }
    }
}
