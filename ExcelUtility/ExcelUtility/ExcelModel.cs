using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    /// <summary>
    /// Populates Excel Class from Excel File in DataTable Format
    /// </summary>
    public class ExcelModel
    {

        DataTable sourceData;
        string comparisonColumn;
        public ExcelModel()
        {

        }
        public ExcelModel(DataTable dataTable)
        {
            this.sourceData = dataTable;
        }

        public DataTable SourceData
        {
            get { return sourceData; }
        }

        public List<string> Columns
        {
            get { return GetColumns(); }
        }

        public int ColIndex { get { return Columns.IndexOf(ComparisonColumn); } }
        /// <summary>
        /// Primary Key , u can change it
        /// </summary>
        public string ComparisonColumn { get { return comparisonColumn; } set { comparisonColumn = value; } }
        private List<string> GetColumns()
        {
          return   sourceData.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToList();
        }

        /// <summary>
        /// Table Intersection
        /// </summary>
        /// <param name="target">The other excel file</param>
        /// <returns></returns>
        public ExcelModel Intersect(ExcelModel target)
        {
            DataTable dt = new DataTable("ResultSet");
            foreach (var item in Columns)
            {
                dt.Columns.Add(item);
            }
            var sourceRows =
    sourceData.Select()
        .Select(dr =>
            dr.ItemArray
                .Select(x => x.ToString())
                .ToArray())
        .ToList();
            var targetRows =
    target.sourceData.Select()
        .Select(dr =>
            dr.ItemArray
                .Select(x => x.ToString())
                .ToArray())
        .ToList();
            foreach (var sourceItem in sourceRows)
            {
                foreach (var targetItem in targetRows)
                {
                    if (targetItem[ColIndex] == sourceItem[ColIndex])
                    {
                        bool isMatch = true;
                        for (int i = 0; i < sourceItem.Length; i++)
                        {
                            if (sourceItem[i] != targetItem[i])
                            {
                                isMatch = false;
                                break;
                            }
                        }

                        if (isMatch)
                        {
                            dt.Rows.Add(sourceItem);
                        }
                        break;
                    }
                }
            }
            return new ExcelModel(dt);
        }

        /// <summary>
        /// Table difference i.e Excel Difference
        /// </summary>
        /// <param name="target"> the other excel file</param>
        /// <returns></returns>
        public ExcelModel Difference(ExcelModel target)
        {
            DataTable dt = new DataTable("ResultSet");
            foreach (var item in Columns)
            {
                if (item != Columns[ColIndex])
                {
                    dt.Columns.Add("Source_" + item.ToString());
                    dt.Columns.Add("Target_" + item.ToString());
                }
                else
                    dt.Columns.Add(item);
            }
            var sourceRows =
    sourceData.Select()
        .Select(dr =>
            dr.ItemArray
                .Select(x => x.ToString())
                .ToArray())
        .ToList();
            var targetRows = 
    target.sourceData.Select()
        .Select(dr =>
            dr.ItemArray
                .Select(x => x.ToString())
                .ToArray())
        .ToList();
            foreach (var sourceItem in sourceRows)
            {
                foreach (var targetItem in targetRows)
                {
                    if(targetItem[ColIndex]==sourceItem[ColIndex])
                    {
                        bool isMatch = true;
                        for (int i = 0; i < sourceItem.Length; i++)
                        {
                            if(i>targetItem.Length || sourceItem[i]!=targetItem[i])
                            {
                                isMatch = false;
                                break;
                            }
                        }

                        if (!isMatch)
                        {
                            List<string> itemArray = new List<string>();
                            itemArray.Add(sourceItem[ColIndex]);
                            for (int i = 0; i < sourceItem.Length; i++)
                            {
                                if (i == ColIndex)
                                    continue;
                                itemArray.Add(sourceItem[i]);
                                itemArray.Add(targetItem[i]);

                            }
                            dt.Rows.Add(itemArray.ToArray());
                        }
                        break;
                    }
                }
            }

            return new ExcelModel(dt);
        }
    }
}
