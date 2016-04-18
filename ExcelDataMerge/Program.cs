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
            setStyles.Add("Sample", CellStyle.NavyFill, CellStyle.Default);
            setStyles.Add("Table 1", CellStyle.LiteBlueFill, CellStyle.Default);

            ExcelExportModel model1 = new ExcelExportModel(dataTable: MyDataStore.GetTableData("Sample",columnCount: 10), sheetName: @"Multi Set Sheet One", exportAs: ExportType.Single);
            ExcelExportModel model2 = new ExcelExportModel(dataSets: dataSets, sheetName: @"Multi Set Sheet Two", exportAs: ExportType.Merged);

            IList<ExcelExportModel> excelList = new List<ExcelExportModel>()
            {
                model1,
                model2
            };

            MultiSetExcelExport.CreateExcelDocument(@"C:\AKK\Excel\Multiple Sets In One.xls", model1, setStyles);
            MultiSetExcelExport.CreateExcelDocument(@"C:\AKK\Excel\Multiple Sheet.xls", excelList, setStyles); 
        }
    }
}
