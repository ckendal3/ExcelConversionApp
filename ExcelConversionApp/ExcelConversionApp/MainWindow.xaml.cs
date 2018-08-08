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
using NPOI;
using System.IO;
using System.ComponentModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.Collections;
using System.Drawing;
using NPOI.XSSF.UserModel;

namespace ExcelConversionApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        NotifyPropertyChange notifyPropertyChange = new NotifyPropertyChange();

        private string fileOpenPath = "None Selected";
        public string FileOpenPath
        {
            get { return fileOpenPath; }
            set
            {
                fileOpenPath = value;
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
            }
        }


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_FileToOpen_Click(object sender, RoutedEventArgs e)
        {
            FindFilePath(ref fileOpenPath, ref fileOpenPathTextBlock);
        }

        private void Button_FileToWrite_Click(object sender, RoutedEventArgs e)
        {
            FindFilePath(ref fileWritePath, ref fileWritePathTextBlock);
        }

        public void FindFilePath(ref string filePath, ref TextBlock textBlock)
        {
            //Open file browser
            Microsoft.Win32.OpenFileDialog fileDialog = new Microsoft.Win32.OpenFileDialog();

            fileDialog.DefaultExt = ".xls";
            fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";

            // Display open file dialog
            Nullable<bool> result = fileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                // Open Document
                filePath = fileDialog.FileName;

                int index = fileOpenPath.LastIndexOf('\\');

                textBlock.Text = FileOpenPath.Substring(index + 1);
                Console.WriteLine("File path: " + fileOpenPath);
            }
        }
    }

    /// <summary>
    /// Read the information from the spreadsheet to import.
    /// </summary>
    public class ExcelReader
    {
        public void ReadWorkBook(string path, List<ContactData> contactList)
        {
            FileStream file = File.OpenRead(path);
            IWorkbook workbook = new XSSFWorkbook(path);

            ContactData tmpContact;
            IRow tmpRow;
            // for every row (contact) in the sheet
            for(int i = 0; i < workbook.GetSheetAt(0).LastRowNum; i++)
            {
                tmpContact = new ContactData();
                tmpRow = workbook.GetSheetAt(0).GetRow(i);

                tmpContact.firstName = tmpRow.GetCell(0).StringCellValue;
                tmpContact.lastName = tmpRow.GetCell(0).StringCellValue;
                tmpContact.emailAddress = tmpRow.GetCell(0).StringCellValue;
                tmpContact.propertyAddress = tmpRow.GetCell(0).StringCellValue;
                tmpContact.phoneNumber = tmpRow.GetCell(0).StringCellValue;
                tmpContact.role = (EMarketRole)tmpRow.GetCell(0).NumericCellValue;


                // manually going to have to count till a better system is found



                var bloop = tmpRow.CopyCell(i, 0).StringCellValue;
                contactList.Add(tmpContact);
            }

                    

        }
    }

    /// <summary>
    /// Creates the new file and the needed data
    /// </summary>
    public class ExcelWriter
    {
        List<ContactData> contactList;

        public void GetAllContactData(List<ContactData> inData)
        {
            contactList = inData;
        }

        public void CreateWorkBook(string path)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");

            IRow tmpRow;
            ContactData contact;

            // For every contact - create new row
            for (int i = 0; i < contactList.Count; i++)
            {
                tmpRow = s1.CreateRow(i);

                contact = contactList[i];

                // Fill sheet with all needed information
                tmpRow.CreateCell(0).SetCellValue(contact.firstName);
                tmpRow.CreateCell(1).SetCellValue(contact.lastName);
                tmpRow.CreateCell(2).SetCellValue(contact.emailAddress);
                tmpRow.CreateCell(3).SetCellValue(contact.propertyAddress);
                tmpRow.CreateCell(4).SetCellValue(contact.phoneNumber);
                tmpRow.CreateCell(5).SetCellValue((int)contact.role);
                
            }

            using (var fs = File.Create(path + "testList.xlsx"))
            {
                workbook.Write(fs);
                fs.Close();
            }  
        }
    }

    public class ContactData
    {
        public ContactData(string inFirstName, string inLastName, string inEmail, string inProperty, string inPhone, int inRole)
        {
            dataElements = 6;

            firstName = inFirstName;
            lastName = inLastName;
            emailAddress = inEmail;
            propertyAddress = inProperty;
            phoneNumber = PhoneNumber(inPhone);
            role = (EMarketRole)inRole;
        }

        public ContactData(string inName, string inEmail, string inProperty, string inPhone, int inRole)
        {
            dataElements = 6;

            string tmp = inName.Trim();
            string[] nameSplit= tmp.Split(new char[] {' '}, 2);

            firstName = nameSplit[0];
            lastName = nameSplit[1];
            emailAddress = inEmail;
            propertyAddress = inProperty;
            phoneNumber = PhoneNumber(inPhone);
            role = (EMarketRole)inRole;
        }

        public ContactData()
        {

        }

        int dataElements;

        public string firstName;
        public string lastName;
        public string emailAddress;
        public string propertyAddress;
        public string phoneNumber;
        public EMarketRole role;

        public static string PhoneNumber(string value)
        {
            value = new System.Text.RegularExpressions.Regex(@"\D")
                .Replace(value, string.Empty);
            value = value.TrimStart('1');
            if (value.Length == 7)
                return Convert.ToInt64(value).ToString("###-####");
            if (value.Length == 10)
                return Convert.ToInt64(value).ToString("###-###-####");
            if (value.Length > 10)
                return Convert.ToInt64(value)
                    .ToString("###-###-#### " + new String('#', (value.Length - 10)));
            return value;
        }
    }

    public enum EMarketRole
    {
        unassigned = 0,
        buyer = 1,
        seller = 2
    };


    public class NotifyPropertyChange : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        internal void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
