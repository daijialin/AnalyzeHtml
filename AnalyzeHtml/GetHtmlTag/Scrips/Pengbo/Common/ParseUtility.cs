using System;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using org.in2bits.MyXls;

namespace Pengbo
{


    public static class ParseUtility
    {
        public static XlsDocument GetInitXlsDoc(string excelName = "newExcel", string author = "admin", string subject = "admin", string sheetName = "Sheet0")
        {
            XlsDocument xls = new XlsDocument();
            xls.FileName = excelName;
            xls.SummaryInformation.Author = author;
            xls.SummaryInformation.Subject = subject;

            xls.Workbook.Worksheets.Add(sheetName);
            return xls;
        }

        /// <summary>
        /// Get Excel Range By Excel Path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
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
