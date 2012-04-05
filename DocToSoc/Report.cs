using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Socrata;
using Socrata.Data.View;

namespace DocToSoc
{
    class Report
    {
        private string shortname;
        private string file;
        private string macroString;
        private bool socrata = false;
        private string socrataAction;
        private string socrataId;
        private int headerRows;
        public Report(string shortname, string file, string macroString, bool socrata = false, string socrataAction = "", string socrataId = "",int headerRows = 0)
        {
            this.shortname = shortname;
            this.file = file;
            this.macroString = macroString;
            this.socrata = socrata;
            this.socrataAction = socrataAction;
            this.socrataId = socrataId;
            this.headerRows = headerRows;
        }

        public void RunMacros()
        {
            if (macroString.Length > 0 && this.file.Length > 0)            
                this.RunMacros(this.file, this.macroString);
        }
                
        public void RunMacros(string file, string macroString)
        {
            {
                string time = DateTime.Now.ToString("yyyyMMddhhmmss");
                string newFile = file.Insert(file.LastIndexOf("."), "_" + time);
                System.IO.File.Copy(file, newFile + "_orig");
                

                System.IO.File.Copy(file, newFile);
                file = newFile;

                string macroPath = System.Configuration.ConfigurationManager.AppSettings["macroPath"];
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //if macroPath(app.config) and xlsb startupPath do not match then copy the macroPath file to the xlsb startup path
                if (macroPath.Length > 0)
                    if (!macroPath.ToLower().Substring(0, macroPath.LastIndexOf("\\")).Equals(xlApp.StartupPath.ToLower()))
                        System.IO.File.Copy(macroPath, xlApp.StartupPath + macroPath.Substring(macroPath.LastIndexOf("\\")));

                string[] macros = macroString.Split('|');
                
                try
                {
                    foreach (string macroName in macros)
                    {
                        xlApp.Run(macroName);
                        this.file = file = file.Substring(0, file.LastIndexOf(".")) + ".csv";
                        xlWorkBook.SaveAs(file, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);//, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlWorkBook.Close(true, misValue, misValue);
                    }
                    

                }

                catch (Exception e)
                {
                    Console.WriteLine("Failed to run report " + file);
                    Console.WriteLine(e.ToString());
                    System.Windows.Forms.MessageBox.Show("Running macro on " + this.shortname + " failed.", "Report Failed", System.Windows.Forms.MessageBoxButtons.OK);
                    throw new Exception("Macro Failed");
                    
                }
                finally
                {
                    //if macroPath(app.config) and xlsb startupPath do not match then delete the macroPath file to the xlsb startup path
                    if (macroPath.Length > 0)
                        if (!macroPath.ToLower().Substring(0, macroPath.LastIndexOf("\\")).Equals(xlApp.StartupPath.ToLower()))
                            System.IO.File.Delete(xlApp.StartupPath + macroPath.Substring(macroPath.LastIndexOf("\\")));

                    
                    xlApp.Quit();
                    Console.WriteLine("Close excel instances in Control Manager. Report: " + this.file);
                }
            }
        }

        public void UploadToSocrata()
        {
            this.UploadToSocrata(this.socrataId,this.file,this.socrataAction);
        }

        public void UploadToSocrata(string socrataId, string file, string socrataAction)
        {
            if (!this.socrata)
                return;

            View v = View.FromId(socrataId);
            bool isPublic = v.IsPublic();
            if(isPublic)
                v = v.WorkingCopy();
            if(socrataAction.Equals("append"))
            {                
                v.Append(file);
            }
            else if (socrataAction.Equals("replace"))
            {
                v.Replace(file,headerRows);
                
            }
            else { }
            if(isPublic)
                v = v.Publish();


        }
        public String MacroString { get { return macroString; } }
    }
}
