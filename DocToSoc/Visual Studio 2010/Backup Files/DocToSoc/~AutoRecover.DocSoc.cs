using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace DocToSoc
{
    class DocSoc
    {
        public DocSoc()
        {
            string file = getFileName();
        }



        public string getFileName()
        {
            string file = "";
            // Show the dialog and get result.
            OpenFileDialog f = new OpenFileDialog();
            if (f.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("Loaded file {0}", f.FileName);
                file = f.FileName;

            }
            return file;
        }

        public void CallMacro(string file)
        {

            Excel.Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file);//, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet
            MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());
            RunMacro(
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

    //        Application.Run "PERSONAL.XLSB!CleanDocket"
    //Application.Run "PERSONAL.XLSB!Create_Upcoming_Docket"
    //Sheets("Upcoming Hearings").Select
    //Sheets("Upcoming Hearings").Move Before:=Sheets(1)



            /*
             * 
             * 
           
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);*/
        }

    }



}
