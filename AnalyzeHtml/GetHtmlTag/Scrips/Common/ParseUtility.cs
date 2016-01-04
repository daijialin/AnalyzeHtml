using System;
using System.Runtime.Remoting.Messaging;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;

namespace GetHtmlTag
{


    public static class ParseUtility
    {
        public static Range GetExcelRang(string path)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[1];
            Range excelRange = worksheet.Cells;
            return excelRange;
        }

        public static HtmlAttribute GetAttribute(HtmlNode sourceNode, string head, string tail)
        {
            HtmlAttribute attribute = null;
            if (sourceNode.Name.Equals(head))
            {
                HtmlAttributeCollection attributeCollection = sourceNode.Attributes;

                if (attributeCollection.Contains(tail))
                {
                    attribute = attributeCollection[tail];
                }
            }
            return attribute;
        }

        public static HtmlNode GetNode(HtmlNode sourceNode, string head, string tail, string tailName)
        {
            HtmlNode newNode = null;
            if (sourceNode.Name.Equals(head))
            {
                HtmlAttributeCollection attributeCollection = sourceNode.Attributes;

                if (attributeCollection.Contains(tail))
                {
                    HtmlAttribute attribute = attributeCollection[tail];
                    string classValue = attribute.Value;

                    if (classValue == tailName)
                    {
                        newNode = attribute.OwnerNode;
                    }
                }
            }
            return newNode;
        }
    }
}
