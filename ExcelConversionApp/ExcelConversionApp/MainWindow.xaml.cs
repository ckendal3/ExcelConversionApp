using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;


/*
 * Spreadsheets are zero based - (A, 1) is (0, 0)
 * 
 * 
 * 
 */
namespace ExcelConversionApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        NotifyPropertyChange notifyPropertyChange = new NotifyPropertyChange();
        List<CellMap> cellMaps = new List<CellMap>();
        
        private string fileOpenPath = "None Selected";
        public string FileOpenPath
        {
            get { return fileOpenPath; }
            set
            {
                fileOpenPath = value;
                Console.WriteLine("File open path: " + value);
                notifyPropertyChange.NotifyPropertyChanged("fileOpenPath");
                
            }
        }

        private string fileWritePath = "None Selected";
        public string FileWritePath
        {
            get { return fileWritePath; }
            set
            {
                fileWritePath = value;
                notifyPropertyChange.NotifyPropertyChanged("fileWritePath");
                Console.WriteLine("File write to path: " + value);
            }
        }


        public MainWindow()
        {
            InitializeComponent();
        }
        
        public void StartParsingProcedure()
        {
            
            if(FileOpenPath == "None Selected" || FileWritePath == "None Selected")
            {
                Console.WriteLine("A file path is not set.");
                return;
            }

            SetMapping();

            ParseFile();

        }


        private void Button_FileToOpen_Click(object sender, RoutedEventArgs e)
        {
            FindFilePath(out string newPath, ref fileOpenPathTextBlock);
            if(!newPath.Equals(""))
            {
                FileOpenPath = newPath;
            }
            
        }

        private void Button_FileToWrite_Click(object sender, RoutedEventArgs e)
        {
            FindFilePath(out string newPath, ref fileWritePathTextBlock);
            if (!newPath.Equals(""))
            {
                FileWritePath = newPath;
            }
        }

        private void Button_StartConversion_Click(object sender, RoutedEventArgs e)
        {
            StartParsingProcedure();
        }

        /// <summary>
        /// Returns the file path for the selected file. 
        /// </summary>
        /// <param name="newPath"></param>
        /// <param name="textBlock"></param>
        public void FindFilePath(out string newPath, ref TextBlock textBlock)
        {
            newPath = "";

            //Open file browser
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();

            fileDialog.DefaultExt = ".xls";
            fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

            // Display open file dialog
            Nullable<bool> result = fileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                // Open Document
                newPath = fileDialog.FileName;

                int index = newPath.LastIndexOf('\\');

                textBlock.Text = newPath.Substring(index + 1);
            }
        }
        
        /// <summary>
        /// This method executes the conversion.
        /// </summary>
        public void ParseFile()
        {      
            ExcelReader reader = new ExcelReader();
            ExcelWriter writer = new ExcelWriter();
            
            // Collected data
            List<NewData> data = reader.ReadWorkBook(FileOpenPath, cellMaps);
            
            // if there is data, write it to the new file with the input name
            if(data.Count > 0)
            {
                if(!fileNameInput.Text.Equals(""))
                {
                    writer.CreateWorkBook(FileWritePath, fileNameInput.Text, data);
                }
                else
                { 
                    writer.CreateWorkBook(FileWritePath, "ConvertedExcelSheet" , data);
                }
                
            }
        }

        public void AddCellMap(CellMap map)
        {
            cellMaps.Add(map);
        }

        public void RemoveCellMap(CellMap map)
        {
            cellMaps.Remove(map);
        }
        
                
        public void SetMapping()
        {
            // When a new item is created, add it to the list cellMapping
            // Create a "prefab" of components the contain 3 input boxes (nameInput, importCellId, newCellId) with a button to submit changes
        }
    }

    public class NotifyPropertyChange : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        internal void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
