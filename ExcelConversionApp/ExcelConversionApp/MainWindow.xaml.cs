﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;


/*
 * Spreadsheets are zero based - (A, 1) is (0, 0)
 * 
 * 
 * 
 */

// TODO: Replace Hardcoded converted file name
// TODO: Offload a lot of code to separate classes
namespace ExcelConversionApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        NotifyPropertyChange notifyPropertyChange = new NotifyPropertyChange();

        ObservableCollection<CellMap> cellMaps = new ObservableCollection<CellMap>();

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

            AddMapControl.button_AddMap.Click += Button_AddMap_Click;

            listview_MappingList.ItemsSource = cellMaps;
        }

        /// <summary>
        /// Adds the mapped cell to the cellmap list and clears the inputs for the next mapping.
        /// </summary>
        private void Button_AddMap_Click(object sender, RoutedEventArgs e)
        {
            AddCellMap(new CellMap(AddMapControl.GetImportId(), AddMapControl.GetExportId(), AddMapControl.GetMapName()));

            AddMapControl.ClearInputs();
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
            Button button = (Button)e.Source;

            Console.WriteLine(button.Content.ToString());

            FindFilePath(out string newPath, ref fileWritePathTextBlock);
            if (!newPath.Equals(""))
            {
                FileWritePath = newPath;
            }
        }

        private void Button_StartConversion_Click(object sender, RoutedEventArgs e)
        {
            if(FilePathsAreSet())
            {
                ParseFile();
            }
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
            // Collected data
            RowData[] data = ExcelReader.ReadWorkBook(FileOpenPath, GetCellMapping(cellMaps)); // get data based on observable cell map list

            if(data == null || data.Length == 0)
            {
                return;
            }

            if(!fileNameInput.Text.Equals(""))
            {
                ExcelWriter.CreateWorkBook(FileWritePath, fileNameInput.Text, data);
            }
            else
            {
                ExcelWriter.CreateWorkBook(FileWritePath, "ConvertedExcelFile", data);
            }

        }

        public void AddCellMap(CellMap map)
        {
            // if data is already there, do nothing
            if(cellMaps.Contains(map))
            {
                return;
            }

            cellMaps.Add(map);
        }

        public void RemoveCellMap(CellMap map)
        {
            cellMaps.Remove(map);
        }

        /// <summary>
        /// Returns a CellMap list rather than an observable collection
        /// </summary>
        /// <param name="observableList"></param>
        /// <returns></returns>
        public CellMap[] GetCellMapping(ObservableCollection<CellMap> observableList)
        {
            int count = observableList.Count;
            CellMap[] tmpArray = new CellMap[count];

            for (int i = 0; i < count; i++)
            {
                tmpArray[i] = observableList[i];
            }

            return tmpArray;
        }

        /// <summary>
        /// This checks if the file paths for reading and writing are set 
        /// </summary>
        /// <returns>Returns true if both paths are set</returns>
        public bool FilePathsAreSet()
        {
            if (FileOpenPath == "None Selected")
            {
                Console.WriteLine("The file to open is not set.");
                return false;
            }

            if(FileWritePath == "None Selected")
            {
                Console.WriteLine("The file location to save to is not set.");
                return false;
            }

            return true;
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