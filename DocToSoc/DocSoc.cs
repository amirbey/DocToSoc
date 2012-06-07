using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Collections;
using System.Xml.XPath;
using System.Xml;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace DocToSoc
{
    class DocSoc
    {

        public DocSoc()
        {
            string reportsFile = System.Configuration.ConfigurationManager.AppSettings["ReportsFile"];
            string macrosFile = System.Configuration.ConfigurationManager.AppSettings["MacrosFile"];
            System.IO.FileInfo macrosFileInfo = new System.IO.FileInfo(macrosFile);
            Hashtable reportsHash = readReportsXML(reportsFile,macrosFileInfo.Name);
            string[] reportsSelected = getSelectedReports(reportsHash);

            Excel.Application xlApp;
            Excel.Workbook macroBook = null;
            
            xlApp = new Excel.Application();
            Console.Write("Excel opened");
           
            try
            {
                macroBook = xlApp.Workbooks.Open(macrosFileInfo.FullName);
                Console.WriteLine("MacroFile opened: " + macrosFileInfo.FullName);
            
                this.processReports(xlApp, reportsHash, reportsSelected);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                System.Windows.Forms.MessageBox.Show("DocToSoc failed while processing reports.", "DocToSoc Failed", System.Windows.Forms.MessageBoxButtons.OK);
            }
            finally
            {
                macroBook.Close();
                Console.WriteLine("MacroWorkBook closed: " + macrosFile);
                xlApp.Quit();
                Console.WriteLine("Excel Closed");
            }
        }

        public void processReports(Excel.Application xlApp, Hashtable reportsHash, string[] selectedReports)
        {
            foreach (string report in selectedReports)
            {
                Report rpt = (Report)reportsHash[report];
                if(rpt.MacroString.Length > 0)
                    rpt.RunMacros(xlApp);
                
                if(rpt.Success)
                    rpt.UploadToSocrata();

            }
        }

        public string[] getSelectedReports(Hashtable reports)
        {
            string[] selectedReports = new string[0];
            ReportsDialog d = new ReportsDialog(reports);
            if (d.ShowDialog() == DialogResult.OK)
            {
                selectedReports = d.ReportsSelectedList;
                Console.WriteLine("Selected Report " + reports.Count /*.Length*/ + " files");
            }
            d.Dispose();
            d.Close();
            return selectedReports;
        }

        public Hashtable readReportsXML(string file, string macroName)
        {
            
            XPathNavigator nav;
            XPathDocument docNav;
            XPathNodeIterator NodeIter;
            String strExpression;
            docNav = new XPathDocument(file);
            nav = docNav.CreateNavigator();

            strExpression = @"/Reports/Report";
            NodeIter = nav.Select(strExpression);
            Hashtable reportsHash = new Hashtable(NodeIter.Count);
            string shortName, name, macro, socrataAction, socrataId, headerRowsString = string.Empty;
            int headerRows = 0;
            bool postToSocrata = false;

            for (int i = 0; NodeIter.MoveNext(); i++)
            {
                XPathNavigator nav2 = NodeIter.Current.Clone();
                shortName = nav2.GetAttribute("Short","");
                name = nav2.GetAttribute("Name", "");
                macro = nav2.GetAttribute("Macro", "");
                if (nav2.GetAttribute("PostToSocrata", "").ToLower().Equals("yes"))
                    postToSocrata = true;
                socrataAction = nav2.GetAttribute("SocrataAction","").ToLower();
                socrataId = nav2.GetAttribute("SocrataId", "").ToLower();
                headerRowsString = nav2.GetAttribute("SocrataHeaderRows", "").Trim();
                if(headerRowsString.Length > 0)
                    headerRows = int.Parse(headerRowsString);
                reportsHash.Add(shortName,new Report(shortName, name, macroName,macro, postToSocrata, socrataAction,socrataId,headerRows));

                Console.WriteLine("Report Added: {" + i + "} => " + shortName);
            }
            return reportsHash;
        }
    }
}