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
        CellMapping cellMap;

        public bool nameIsSingleCell = true;
        
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

        private void Button_IsNameCombined_Click(object sender, RoutedEventArgs e)
        {
            if(nameIsSingleCell == true)
            {
                Button_IsNameCombined.Content = "No";
                nameIsSingleCell = false;
            }
            else
            {
                Button_IsNameCombined.Content = "Yes";
                nameIsSingleCell = true;
            }
        }

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
        
        public void ParseFile()
        {      
            ExcelReader reader = new ExcelReader();
            ExcelWriter writer = new ExcelWriter();
            
            // Collected data
            List<ContactData> contacts = reader.ReadWorkBook(FileOpenPath, cellMap, nameIsSingleCell);
            
            // if there is data, write it to the new file
            if(contacts.Count > 0)
            {
                writer.CreateWorkBook(FileWritePath, contacts);
            }
        
        }
        
                
        public void SetMapping()
        {
            List<int> mappingList = new List<int>();

            mappingList.Add(Convert.ToInt32(nameSlider.Value));
            mappingList.Add(Convert.ToInt32(firstNameSlider.Value));
            mappingList.Add(Convert.ToInt32(lastNameSlider.Value));
            mappingList.Add(Convert.ToInt32(emailSlider.Value));
            mappingList.Add(Convert.ToInt32(propertySlider.Value));
            mappingList.Add(Convert.ToInt32(phoneSlider.Value));
            mappingList.Add(Convert.ToInt32(roleSlider.Value));


            if (nameIsSingleCell)
            {
                cellMap = new CellMapping(mappingList[0], mappingList[3], mappingList[4], mappingList[5], mappingList[6]); // name is combined
            }
            else
            {
                cellMap = new CellMapping(mappingList[1], mappingList[2], mappingList[3], mappingList[4], mappingList[5], mappingList[6]); // first and last name already separated
            }
        }
    }

    /// <summary>
    /// Read the information from the spreadsheet to import.
    /// </summary>
    public class ExcelReader
    {
        public List<ContactData> ReadWorkBook(string path, CellMapping map, bool nameIsSingleCell)
        {
            List<ContactData> contactList = new List<ContactData>();
            
            FileStream file = File.OpenRead(path);
            IWorkbook workbook = new XSSFWorkbook(path);
            ISheet sheet = workbook.GetSheetAt(0);

            ContactData tmpContact;
            IRow tmpRow;
            // for every row (contact) in the sheet
            for(int i = 0; i < workbook.GetSheetAt(0).LastRowNum + 1; i++)
            {
                tmpRow = sheet.GetRow(i);

                // Get data from specified cells
                // User input can determine which cell to get the data from
                
                if(nameIsSingleCell)
                {
                    // Create contact constructor   
                    tmpContact = new ContactData(tmpRow.GetCell(0).StringCellValue, // combined name
                                                tmpRow.GetCell(1).StringCellValue,
                                                tmpRow.GetCell(2).StringCellValue,
                                                tmpRow.GetCell(3).NumericCellValue.ToString(),
                                                //Convert.ToInt32(tmpRow.GetCell(map.phoneIndex).StringCellValue),
                                                Convert.ToInt32(tmpRow.GetCell(4).NumericCellValue));
                }
                else
                {
                    tmpContact = new ContactData(tmpRow.GetCell(map.firstNameIndex).StringCellValue, // first name
                                                 tmpRow.GetCell(map.lastNameIndex).StringCellValue, //last name
                                                 tmpRow.GetCell(map.emailIndex).StringCellValue,
                                                 tmpRow.GetCell(map.propertyIndex).StringCellValue,
                                                 tmpRow.GetCell(map.phoneIndex).StringCellValue,
                                                 Convert.ToInt32(tmpRow.GetCell(map.roleIndex).NumericCellValue));
                }
                contactList.Add(tmpContact);
            }

            return contactList;
        }
      
    }

    /// <summary>
    /// Creates the new file and populates it with the directed information.
    /// </summary>
    public class ExcelWriter
    {
        public void CreateWorkBook(string path, List<ContactData> inData)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");

            IRow tmpRow;
            ContactData contact;

            // For every contact - create new row
            for (int i = 0; i < inData.Count; i++)
            {
                tmpRow = s1.CreateRow(i);

                contact = inData[i];

                // Fill sheet with all needed information
                tmpRow.CreateCell(0).SetCellValue(contact.firstName);
                tmpRow.CreateCell(1).SetCellValue(contact.lastName);
                tmpRow.CreateCell(2).SetCellValue(contact.emailAddress);
                tmpRow.CreateCell(3).SetCellValue(contact.propertyAddress);
                tmpRow.CreateCell(4).SetCellValue(contact.phoneNumber);
                tmpRow.CreateCell(5).SetCellValue((int)contact.role);
                
            }

            int index = path.LastIndexOf('\\');

            path = path.Substring(0, index + 1);

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
            firstName = inFirstName;
            lastName = inLastName;
            emailAddress = inEmail;
            propertyAddress = inProperty;
            phoneNumber = PhoneNumber(inPhone);
            role = (EMarketRole)inRole;
        }

        public ContactData(string inName, string inEmail, string inProperty, string inPhone, int inRole)
        {
            string tmp = inName.Trim();
            string[] nameSplit= tmp.Split(new char[] {' '}, 2);

            firstName = nameSplit[0];
            lastName = nameSplit[1];
            emailAddress = inEmail;
            propertyAddress = inProperty;
            phoneNumber = PhoneNumber(inPhone);
            role = (EMarketRole)inRole;
        }

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
    
    public class CellMapping
    {
        
        public CellMapping(int nameDex, int emailDex, int propertyDex, int phoneDex, int roleDex)
        {
            nameIndex = nameDex;
            emailIndex = emailDex;
            propertyIndex = propertyDex;
            phoneIndex = phoneDex;
            roleIndex = roleDex;

            Console.WriteLine("Name needs splitting: " + nameIndex.ToString() + emailIndex.ToString() + propertyIndex.ToString() + phoneIndex.ToString() + roleIndex.ToString());
        }
        
        public CellMapping(int firstNameDex, int lastNameDex, int emailDex, int propertyDex, int phoneDex, int roleDex)
        {
            firstNameIndex = firstNameDex;
            lastNameIndex = lastNameDex;
            emailIndex = emailDex;
            propertyIndex = propertyDex;
            phoneIndex = phoneDex;
            roleIndex = roleDex;

            Console.WriteLine("Name Already split:" + firstNameIndex.ToString() + lastNameIndex.ToString() + emailIndex.ToString() + propertyIndex.ToString() + phoneIndex.ToString() + roleIndex.ToString());
        }
        
        
        public int nameIndex;
        
        public int firstNameIndex;
        public int lastNameIndex;
        public int emailIndex;
        public int propertyIndex;
        public int phoneIndex;
        public int roleIndex;
        
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
