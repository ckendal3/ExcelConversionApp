using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelConversionApp
{
    /// <summary>
    /// Interaction logic for AddMapControl.xaml
    /// </summary>
    public partial class AddMapControl : UserControl
    {
        public AddMapControl()
        {
            InitializeComponent();
        }


        public int GetImportId()
        {
            return Convert.ToInt32(input_ImportID.Text);
        }

        public int GetExportId()
        {
            return Convert.ToInt32(input_ExportID.Text);
        }

        public string GetMapName()
        {
            return input_MapName.Text;
        }

        public void ClearInputs()
        {
            input_ImportID.Clear();
            input_ExportID.Clear();
            input_MapName.Clear();
        }
    }
}
