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
    /// Interaction logic for MapComponent.xaml
    /// </summary>
    public partial class MapComponent : UserControl
    {
        public MapComponent()
        {
            InitializeComponent();
        }

        public void SetImportID(int newValue)
        {
            textblock_ImportID.Text = newValue.ToString();
        }

        public void SetExportId(int newValue)
        {
            textblock_exportID.Text = newValue.ToString();
        }

        public void SetMapName(string newValue)
        {
            textblock_MapName.Text = newValue;
        }

        // Add remove map function
    }
}
