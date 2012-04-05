using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace DocToSoc
{
    public partial class ReportsDialog : Form
    {
        public string[] reportsSelectedList = null;
        public ReportsDialog(Hashtable reports)//string[] reports)
        {
            InitializeComponent();
            addReports(reports);
        }

        public void addReports(Hashtable reports)//string[] reports)
        {
            this.reportsListBox.Items.AddRange(new ArrayList(reports.Keys).ToArray());
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            int numSelectedReports = this.reportsListBox.SelectedItems.Count;
            reportsSelectedList = new string[numSelectedReports];
            for (int i = 0; i < numSelectedReports; i++)
                reportsSelectedList[i] = (string)this.reportsListBox.SelectedItems[i];
        }
        public string[] ReportsSelectedList{
            get { return this.reportsSelectedList; }
        }
    }
}
