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

namespace SheetSplit
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            SplitSheets();
        }

        static void SplitSheets()
        {
            MessageBox.Show("先试一下");

        }

        static void SplitAndMail(string FileName)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;

            FileInfo fi = new FileInfo(FileName);
            string FullFileName = fi.FullName.ToString();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(FullFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            }
            catch
            {
                Console.WriteLine("You must choose a file");
                return;
            }
            xlWorkBook = xlApp.Workbooks.Open(FullFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Range rng = xlApp.get_Range("A1");
            //int index = 0;
            string liteFileName = FileName.Replace(".xlsx", "");
            foreach (Microsoft.Office.Interop.Excel.Worksheet displayWorksheet in xlApp.Worksheets)
            {

                string root = fi.Directory.ToString();
                string sheetName = displayWorksheet.Name.ToString();
                string SaveFileName = root + "\\" + liteFileName + "." + sheetName + ".xlsx";
                //Microsoft.Office.Interop.Excel.Application NewxlApp;
                //NewxlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook NewWorkbook;
                NewWorkbook = xlApp.Workbooks.Add();

                Microsoft.Office.Interop.Excel.Worksheet NewWorkSheet;
                NewWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)NewWorkbook.Worksheets.Add();

                displayWorksheet.Copy(NewWorkSheet);
                NewWorkbook.SaveAs(SaveFileName);
                NewWorkbook.Save();
                NewWorkbook.Close();
            }
        }
    }
}
