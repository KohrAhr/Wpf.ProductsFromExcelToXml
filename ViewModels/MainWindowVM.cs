using Core.Major;
using ExcelToXML.Core;
using ExcelToXML.Functions;
using ExcelToXML.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using ExcelToXML.Core.Major;
using System.Windows.Threading;
using Microsoft.Win32;

namespace ExcelToXML.ViewModel
{
    public class MainWindowVM : PropertyChangedNotification
    {
        private List<ProductBlock> productsBlock = new List<ProductBlock>();

        /// <summary>
        ///     File for proceed
        /// </summary>
        public string FileName
        {
            get => GetValue(() => FileName);
            set => SetValue(() => FileName, value);
        }

        /// <summary>
        ///     Activity log
        /// </summary>
        public string Log
        {
            get { return GetValue(() => Log); }
            set { SetValue(() => Log, value); }
        }

        /// <summary>
        ///     Indicate that we could run Process
        /// </summary>
        public bool ReadyToRun
        {
            get { return GetValue(() => ReadyToRun); }
            set { SetValue(() => ReadyToRun, value); }
        }

        /// <summary>
        ///     Constructor
        /// </summary>
        public MainWindowVM()
        {
            InitCommands();

            InitData();
        }

        private void InitCommands()
        {
            NewXmlCommand = new RelayCommand(newXmlCommand);
            ClearLogCommand = new RelayCommand(clearLogCommand);
            SelectExcelFileCommand = new RelayCommand(selectExcelFileCommand);
        }

        #region Commands
        public ICommand NewXmlCommand { get; set; }
        public ICommand ClearLogCommand { get; set; }
        public ICommand SelectExcelFileCommand { get; set; }
        #endregion Commands

        private void InitData()
        { 
            FileName = @"E:\Temp\GitHubUnsorted\Wpf.ProductsFromExcelToXml\test.xlsx";
            Log += DateTime.Now.ToString() + "\nREADY!\n";
            ReadyToRun = true;
        }

        private void selectExcelFileCommand(Object o)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = Environment.CurrentDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                FileName = openFileDialog.FileName;
            }
        }

        private void clearLogCommand(Object o)
        {
            Log = "";
        }

        private void newXmlCommand(Object o)
        {
            ReadyToRun = false;

            new Task(() =>
            {
                Log += DateTime.Now.ToString() + "\nFile: " + FileName + "\n";

                ExcelFile excelFile = null;
                try
                {
                    productsBlock.Clear();

                    excelFile = ExcelFunctions.OpenExcelFile(FileName);

                    if (excelFile == null)
                    {
                        Log += "Cannot open requested Excel file";
                        return;
                    }

                    // for each Sheet
                    Log += "Analazying worksheets";
                    for (int x = 0; x++ < excelFile.sheet.Worksheets.Count;)
                    {
                        excelFile.worksheet = excelFile.sheet.Worksheets[x];

                        // proceed worksheet analyze
                        productsBlock.AddRange(CoreFunctions.AnalyzeWorksheet(excelFile.worksheet));
                        Log += ".";
                    }

                    Log += "\nTotal blocks detected: " + productsBlock.Count.ToString() + "\n";

                    // Now I know how much products we have in file and where is it located
                    // Go gather information and build XML
                    Log += "Parsing products blocks";
                    string xmlAsText = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<products>";
                    foreach (ProductBlock productBlock in productsBlock)
                    {
                        xmlAsText += CoreFunctions.GatherProductInformation(productBlock, excelFile);
                        Log += ".";
                    }
                    xmlAsText += "\n</products>";

                    System.IO.File.WriteAllText(FileName + ".xml", xmlAsText, Encoding.UTF8);
                }
                finally
                {
                    ExcelFunctions.CloseExcelFile(excelFile);
                }

                Log += "\n" + DateTime.Now.ToString() + "\nDone!\n";

                ReadyToRun = true;
            }).Start();
        }
    }
}
