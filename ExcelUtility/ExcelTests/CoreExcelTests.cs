using System;
using ExcelUtility;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelTests
{
    [TestClass]
    public class CoreExcelTests
    {

        /// <summary>
        /// For now let's test the code
        /// </summary>
        [TestMethod]
        
        public void IntegrationTest()
         {
            string sourceExcel = @"C:\Users\dhara\Downloads\Comparison\Source.xlsx";
            string targetExcel = @"C:\Users\dhara\Downloads\Comparison\Target.xlsx";
            //source excel
            var sourceDataTable = Utils.ExcelToDataTable(sourceExcel,".xlsx");

            //target excel
            var targetDataTable = Utils.ExcelToDataTable(targetExcel,".xlsx");

            //Creating models from excels
            var sourceModel = new ExcelModel(sourceDataTable);
            var targetModel = new ExcelModel(targetDataTable);
            sourceModel.ComparisonColumn = sourceModel.Columns[0];//setting primary key here
            targetModel.ComparisonColumn = sourceModel.ComparisonColumn;
            var match = sourceModel.Intersect(targetModel);
            match.ComparisonColumn = sourceModel.ComparisonColumn;
            var nonMatch = sourceModel.Difference(targetModel);
            var savedMatch = Utils.SaveDataTableExcel(match.SourceData, "match.xlsx");

            // u get the idea rit?.. i am reading excel, converting to Datatable and querying the table
            //if u want more options then u have to do
            // Import excel , convert to DataTable, then store in Sqlite.Once u save to sqlite then u can connect to SQLITE using 
            //one of the SQL Providers like SQLIteAdapter, then u can query the db like u do a normal db like this
            var savedNonMatch = Utils.SaveDataTableExcel(nonMatch.SourceData, "nonMatch.xlsx");
        }
    }
}
