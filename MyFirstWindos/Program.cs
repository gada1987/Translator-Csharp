using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Cells;
using System.Xml;
using System.Xml.XPath;
using System.IO;
using System.Web;



namespace MyFirstWindos
{
  public static class Program
    {
        [System.STAThreadAttribute()]
        public static void Main()
        {
          Application.EnableVisualStyles();
          Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new login());
            Application.Run(new Form1());
           
        }
    }
}


        
    





       

