using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.Xml;

namespace MyFirstWindos
{
   public class Compare_D_R
    {
      
      public void compareD_R(string excelFilePath, string def,string alt)
        {
            //läs in Excel-filen
            Workbook wb = new Workbook();
            wb.Open(excelFilePath);
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;
            int last_row = ws.Cells.MaxRow;

            //läs in den svenska version
            XmlDocument doc1 = new XmlDocument();
            doc1.Load(def);
            //läs in Alternative file
            XmlDocument doc2 = new XmlDocument();
            doc2.Load(alt);
            XmlNode XNode = doc2.SelectSingleNode("/root");
            //jämför filerna
            var children1 = doc1.SelectNodes("root/data");
            for (int i = 1; i <= last_row; i++)
            {
                bool cond = false;
                for (var d = 0; d < children1.Count; d++)
                {

                    var child1 = children1[d];

                    {
                        if (cells[i, 0].StringValue.ToString() == child1.Attributes["name"].Value)
                        {
                            //var childclone = child1.CloneNode(true);
                            child1.ChildNodes[1].InnerText = cells[i, 1].StringValue.ToString();
                            XNode.AppendChild(doc2.ImportNode(child1, true));
                        }

                        else
                        {
                            cond = true;
                        }

                    }
                }
            }
          
            doc2.Save(alt);
            string DecodedString = System.Web.HttpUtility.HtmlDecode(alt.ToString());

        }
    }
}
    

    

