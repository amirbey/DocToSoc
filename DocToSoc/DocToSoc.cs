using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DocToSoc
{
    class DocToSoc
    {
        public DocToSoc(){
            string file = "";
        }

        public string getFileName(){
           string file = "";
            // Show the dialog and get result.
	    DialogResult result = openFileDialog1.ShowDialog();
	    if (result == DialogResult.OK) // Test result.
	    {
            result.
	    }
	    Console.WriteLine(result); // <-- For debugging use only. 
        }
    }
}
