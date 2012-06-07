using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Socrata;
using Socrata.Data.View;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocToSoc
{
    class Report
    {
        private string shortname;
        private string file;
        private string macroString;
        private string macroFile;
        private bool socrata = false;
        private string socrataAction;
        private string socrataId;
        private int headerRows;
        private bool success = true;
        public Report(string shortname, string file, string macroFile, string macroString, bool socrata = false, string socrataAction = "", string socrataId = "",int headerRows = 0)
        {
            this.shortname = shortname;
            this.file = file;
            this.macroFile = macroFile;
            this.macroString = macroString;
            this.socrata = socrata;
            this.socrataAction = socrataAction;
            this.socrataId = socrataId;
            this.headerRows = headerRows;
        }

        public void RunMacros(Excel.Application xlApp)
        {
            if (macroString.Length > 0 && this.file.Length > 0)            
                this.RunMacros(xlApp, this.file, this.macroFile, this.macroString);
        }
                
        public void RunMacros(Excel.Application xlApp, string file, string macroFileName, string macroString)
        {
            {
                string time = DateTime.Now.ToString("yyyyMMddhhmmss");
                string newFile = file.Insert(file.LastIndexOf("."), "_" + time);
                System.IO.File.Copy(file, newFile + "_orig");
                

                System.IO.File.Copy(file, newFile);
                file = newFile;

                Excel.Workbook xlWorkBook;
            
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                string[] macros = macroString.Split('|');
                //macroBook = xlApp.Workbooks.Open(macroFile);
                xlWorkBook.Activate();
                
                try
                {
                    foreach (string macroName in macros)
                        xlApp.Run(macroFileName + "!" + macroName);
                
                    this.file = file = file.Substring(0, file.LastIndexOf(".")) + ".csv";
                    xlWorkBook.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);//, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //xlWorkBook.Close(true, misValue, misValue);
                    
                }

                catch (Exception e)
                {
                    this.success = false;
                    Console.WriteLine("Failed to run report " + file);
                    Console.WriteLine(e.ToString());
                    System.Windows.Forms.MessageBox.Show("Running macro on " + this.shortname + " failed.", "Report Failed", System.Windows.Forms.MessageBoxButtons.OK);
                    //throw new Exception("Macro Failed");
                }
                finally
                {
                    xlWorkBook.Close(false);
                    Console.WriteLine("Excel Workbook Closed. Report: " + this.file);
                }
            }
        }

        public void UploadToSocrata()
        {
            this.UploadToSocrata(this.socrataId,this.file,this.socrataAction);
        }

        public void UploadToSocrata(string socrataId, string file, string socrataAction)
        {
            try
            {
                if (!this.socrata)
                    return;

                socrataAction = socrataAction.ToLower();
                View v = View.FromId(socrataId);
                bool isPublic = v.IsPublic();
                if (isPublic)
                    v = v.WorkingCopy();
                if (socrataAction.Equals("append"))
                {
                    v.Append(file);
                }
                else if (socrataAction.Equals("replace"))
                {
                    v.Replace(file, headerRows);

                }
                else { }
                if (isPublic)
                    v = v.Publish();

            }
            catch (Exception e)
            {
                this.success = false;
                Console.WriteLine("Failed to upload report to Socrata" + file);
                Console.WriteLine(e.ToString());
                System.Windows.Forms.MessageBox.Show("Uploading file " + this.shortname + " to Socrata failed.", "Socrata Upload Failed", System.Windows.Forms.MessageBoxButtons.OK);
                //throw new Exception("Macro Failed");
            }
        }
        public String MacroString { get { return macroString; } }
        public bool Success { get { return success; } }
    }
}
