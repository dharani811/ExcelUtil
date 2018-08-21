using ExcelUtility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UI.ViewModels
{
    public class ViewGridController:NotifyUI
    {
        private ExcelModel model;
        private bool isPrimaryKeyEnabled;
        private int selectedKeyIndex=0;
        
        public ViewGridController(string filename,bool keySelectionEnabled)
        {
            isPrimaryKeyEnabled = keySelectionEnabled;
            var data = Utils.ExcelToDataTable(filename,".xlsx");
            model = new ExcelModel(data);
            SelectedKeyIndex = 0;
        }
        public ExcelModel Model { get { return model; } set { model = value; } }

        public bool IsPrimaryKeyEnabled {
            get => isPrimaryKeyEnabled;
            set => isPrimaryKeyEnabled = value; }
        public int SelectedKeyIndex { get { return selectedKeyIndex; }
            set {
                
                    selectedKeyIndex = value; UpdateUI("SelectedKeyIndex");
                    model.ComparisonColumn = model.Columns[selectedKeyIndex];
                
            } }
    }
}
