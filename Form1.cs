using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using System.Linq;

namespace FolderListExporter
{
    public partial class FolderListExporter : Form
    {
        public FolderListExporter()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialog = new FolderBrowserDialog();
            var returnValue = FolderBrowserDialog.ShowDialog();
            if (returnValue == DialogResult.OK) {
                label2.Text=FolderBrowserDialog.SelectedPath;
                button2.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            DirectoryInfo di = new DirectoryInfo(label2.Text);
            

            FileInfo[] fileNames = di.GetFiles();

            ArrayList filenameList = new ArrayList();

            
            foreach (FileInfo fileName in fileNames) {
                var fileNameHolders = new FileNameHolder();
                fileNameHolders.name = fileName.Name;
                fileNameHolders.size = fileName.Length;
                filenameList.Add(fileNameHolders);
            }


            foreach (DirectoryInfo fileName in di.GetDirectories())
            {
                var fileNameHolders = new FileNameHolder();
                fileNameHolders.name = fileName.Name;
                fileNameHolders.size = GetDirectorySize(fileName);
                filenameList.Add(fileNameHolders);
            }

            filenameList.Sort();

            exportFile(filenameList.ToArray(typeof(FileNameHolder)) as FileNameHolder[], di.Name);
            linkLabel1.Text = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)+"\\"+ di.Name+".xlsx";
        }

        private static long GetDirectorySize(DirectoryInfo di)
        {
            
            return di.EnumerateFiles("*.*", SearchOption.AllDirectories).Sum(fi => fi.Length);
        }

        private void exportFile(FileNameHolder[] files, string filename)
        {

            //export to xlsx
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = false;
                oXL.DisplayAlerts = false;


                //Get a new workbook.
                oWB = oXL.Workbooks.Add("");
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "File Name";
                oSheet.Cells[1, 2] = "Size";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[files.Length, 2];

                int count = 0;
                foreach(FileNameHolder s in files){
                    saNames[count, 0] = s.name;
                    saNames[count, 1] = ConvertSizeToText(s.size);
                    count++;
                }
                
                

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B"+(files.Length+1)).Value2 = saNames;

                //AutoFit columns A:D.
                oRng=oSheet.get_Range("A1", "B1");
                oRng.EntireColumn.AutoFit();

                oXL.Visible = false;
                oXL.UserControl = false;

               
                var exportFilename = filename + ".xlsx";
                oWB.SaveAs(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + exportFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                    false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                oWB.Close();
                oXL.Quit();
                
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
            
        }

        private string ConvertSizeToText(long size) {

            int level = 0;

            while (size > 1024)
            {    
                level++;
                size = size / 1024;
            }

            string suffix = "b";
            switch (level) {
                case 1:
                    suffix = "KB";
                    break;
                case 2:
                    suffix = "MB";
                    break;
                case 3:
                    suffix = "GB";
                    break;
                case 4:
                    suffix = "TB";
                    break;
                case 5:
                    suffix = "PB";
                    break;
            }

            return size+" "+suffix;
        }

        private void link_clicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabel1.Text);

        }
    }

    class FileNameHolder : IComparable {

        

        public string name { get; set; }

        public long size { get; set;  }

        public int CompareTo(object obj)
        {
            return name.CompareTo(((FileNameHolder)obj).name);
        }
    }
}
