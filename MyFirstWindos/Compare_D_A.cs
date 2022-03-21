using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Aspose.Cells;

namespace MyFirstWindos
{
    public class Compare_D_A
    {
        string excelFilePath;
        public void CompareD_A(string defaultFilePath, string alt,string res)
        {

            XmlDocument doc1 = new XmlDocument();
            doc1.Load(defaultFilePath);

            XmlDocument doc2 = new XmlDocument();
            doc2.Load(alt);

            var result = new XmlDocument();
            result.LoadXml("<root></root>");

            // XmlElement root = doc2.DocumentElement;

            var children1 = doc1.SelectNodes("root/data");
            var children2 = doc2.SelectNodes("root/data");

            XmlNode XNode = result.SelectSingleNode("/root");


            for (var d = 0; d < children1.Count; d++)
            {
                var exists = false;
                var child1 = children1[d];
                for (var x = 0; x < children2.Count; x++)

                {
                    var child2 = children2[x];

                    if (child1.Attributes["name"].Value == child2.Attributes["name"].Value)
                    {

                        exists = true;

                    }
                }
                if (!exists)
                {

                    XNode.AppendChild(result.ImportNode(child1, true));

                }
            }
            result.Save(res);
            excelFilePath = defaultFilePath.Substring(0, defaultFilePath.LastIndexOf("\\") + 1);
            string DecodedString = System.Web.HttpUtility.HtmlDecode(res.ToString());

            XmlNodeList Xlist = result.SelectNodes("/root/data");

            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            ws = wb.Worksheets[0];
            ws.Cells[0, 0].PutValue("name");
            ws.Cells[0, 1].PutValue("Value");

            var i = 1;

            foreach (XmlNode xn in Xlist)
            {
                ws.Cells[i, 0].PutValue(xn.Attributes["name"].Value);
                ws.Cells[i, 1].PutValue(xn.ChildNodes[1].InnerText);
                i++;
            }
            wb.Save(excelFilePath + "sv-SE_to_en-GB.xlsx");



        }
    }
}

