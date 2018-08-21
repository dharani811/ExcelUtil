using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace UI.ViewModels
{
    public class NotifyUI : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public void UpdateUI(string propertyName)
        {
            if(PropertyChanged!=null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
