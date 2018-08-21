using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelUtility;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace UI.ViewModels
{
    public class MainController:NotifyUI
    {
        private ViewGridController sourceModel;
        private ViewGridController targetModel;
        private System.Data.DataTable resultModel;
        private string sourceFilePath;
        private string targetFilePath;

        private int opSelectedIndex=-1;
        private ICommand sourceSelectCommand;
        private ICommand sourceTargetCommand;
        private ICommand exportCommand;
        public MainController()
        {
            opSelectedIndex = 0;
            sourceFilePath = "";
            targetFilePath = "";
            sourceSelectCommand = new RelayCommand((p) => ChooseSource());
            sourceTargetCommand = new RelayCommand((p) => ChooseTarget());
            exportCommand = new RelayCommand((p) => Export());
        }

        private void Export()
        {
            if (ResultModel != null)
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Title = "Save Result";
               saveDialog.Filter = "Excel 97-2003 WorkBook|*.xls|Excel WorkBook|*.xlsx|All Excel Files|*.xls;*.xlsx|All Files|*.*";
                if ((bool)saveDialog.ShowDialog())
                {
                    using (var workbook = SpreadsheetDocument.Create(saveDialog.FileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
                    {
                        var workbookPart = workbook.AddWorkbookPart();

                        workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                        workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                        var table = ResultModel;

                            var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                            var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                            sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                            DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                            string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                            uint sheetId = 1;
                            if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                            {
                                sheetId =
                                    sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                            }

                            DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                            sheets.Append(sheet);

                            DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                            List<String> columns = new List<string>();
                            foreach (System.Data.DataColumn column in table.Columns)
                            {
                                columns.Add(column.ColumnName);

                                DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                                headerRow.AppendChild(cell);
                            }


                            sheetData.AppendChild(headerRow);

                            foreach (System.Data.DataRow dsrow in table.Rows)
                            {
                                DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                                foreach (String col in columns)
                                {
                                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                                    newRow.AppendChild(cell);
                                }

                                sheetData.AppendChild(newRow);
                            }

                        
                    }
                    MessageBox.Show("File Saved Successfully");
                }
            }
        }

        public ViewGridController SourceModel { get { return sourceModel; }

            set { sourceModel = value;UpdateUI("SourceModel"); }
        }
        public ViewGridController TargetModel { get { return targetModel; }

            set { targetModel = value; UpdateUI("TargetModel"); }

        }

        public DataTable ResultModel
        {

            get { return resultModel; }
            set { resultModel = value;

                UpdateUI("ResultModel");
            }
        }

        public List<string> Operations
        {
           get { return GetOps(); }
        }

        public int OpSelectedIndex { get {return  opSelectedIndex; }
            set {
                if (value != opSelectedIndex)
                {
                    opSelectedIndex = value;
                    PerformOperation();
                    UpdateUI("OpSelectedIndex");
                }
            }
        }

        public string SourceFilePath { get { return sourceFilePath; } set { sourceFilePath = value;UpdateUI("SourceFilePath"); } }
        public string TargetFilePath { get { return targetFilePath; } set { targetFilePath = value; UpdateUI("TargetFilePath"); } }

        public ICommand SourceSelectCommand { get { return sourceSelectCommand; } set { sourceSelectCommand = value; } }
        public ICommand SourceTargetCommand { get { return sourceTargetCommand; } set { sourceTargetCommand = value; } }
        public ICommand ExportCommand { get { return exportCommand; } set { exportCommand = value; } }

        private void PerformOperation()
        {
            sourceModel.SelectedKeyIndex = sourceModel.SelectedKeyIndex;
            targetModel.SelectedKeyIndex = sourceModel.SelectedKeyIndex;
            DataTable table =  null;
            if(opSelectedIndex==1)
            {
                 var result = sourceModel.Model.Intersect(targetModel.Model);
                if (result != null)
                    table = result.SourceData;
            }
            else if(opSelectedIndex==2)
            {
                 var result = sourceModel.Model.Difference(targetModel.Model);
                if (result != null)
                    table = result.SourceData;

            }
            ResultModel = table;
        }

        private List<string> GetOps()
        {
            return new List<string>() { "None","Matching","Non-Matching"};
        }

        private void ChooseSource ()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
           bool dialogResult =(bool) openDialog.ShowDialog();
            openDialog.Title = "ChooseSource";
            if(dialogResult)
            {
                SourceModel = new ViewGridController(openDialog.FileName,true);
                SourceFilePath = openDialog.FileName;
            }
        }

        private void ChooseTarget()
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            bool dialogResult = (bool)openDialog.ShowDialog();
            openDialog.Title = "ChooseTarget";
            if (dialogResult)
            {
                TargetModel = new ViewGridController(openDialog.FileName,false);
                TargetFilePath = openDialog.FileName;
            }
        }
    }
}
