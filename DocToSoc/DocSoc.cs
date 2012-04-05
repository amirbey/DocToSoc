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

namespace DocToSoc
{
    class DocSoc
    {

        public DocSoc()
        {
            string reportsFile = System.Configuration.ConfigurationManager.AppSettings["ReportsFile"];
            Hashtable reportsHash = readReportsXML(reportsFile);
            string[] reportsSelected = getSelectedReports(reportsHash);
        
            this.processReports(reportsHash,reportsSelected);
        }

        public void processReports(Hashtable reportsHash, string[] selectedReports)
        {
            foreach (string report in selectedReports)
            {
                Report rpt = (Report)reportsHash[report];
                if(rpt.MacroString.Length > 0)
                    rpt.RunMacros();
                
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

        public Hashtable readReportsXML(string file)
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
                reportsHash.Add(shortName,new Report(shortName, name, macro, postToSocrata, socrataAction,socrataId,headerRows));

                Console.WriteLine("Report Added: {" + i + "} => " + shortName);
            }
            return reportsHash;
        }
    }
}