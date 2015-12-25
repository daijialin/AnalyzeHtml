using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Reflection;
using System.Data;
using org.in2bits.MyXls;
using Microsoft.Office.Interop.Excel;


namespace GetHtmlTag
{

    class MainClass
    {
        private static Dictionary<string, string> m_resultDic = new Dictionary<string, string>();
        private static List<ItemData> iconDataList = new List<ItemData>();


        static void Main(string[] args)
        {
            ParseHtml parseHtml = new ParseHtml();
            parseHtml.ReadXml();
            //DoHtml();
        }
        private static void DoHtml()
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(HtmlConst.m_url);
            HtmlNode mainNode = doc.GetElementbyId(HtmlConst.kRootElement);

            InitResultDic(mainNode);
            InitIconDataList();
            ExportExcel(HtmlConst.kExportExcelName, iconDataList);
            DownloadImage(iconDataList);
            //Console.WriteLine("ok");
        }

        private static void InitIconDataList()
        {
       
            int index = 0;
            foreach (KeyValuePair<string, string> item in m_resultDic)
            {
                ItemData td = new ItemData();
                td.m_id = index;
                td.m_heroName = item.Key;
                td.m_iconPath = item.Value;
                index++;
                iconDataList.Add(td);
            }
        }

        private static void InitResultDic(HtmlNode mainNode)
        {
            if (mainNode.HasChildNodes)
            {
                for (int i = 0; i < mainNode.ChildNodes.Count; i++)
                {
                    HtmlNode tempNode = mainNode.ChildNodes[i];
                    if (tempNode.Name.Equals("img"))
                    {
                        HtmlAttributeCollection attributeCollection = tempNode.Attributes;
                        string srcValue = string.Empty;
                        string alt = string.Empty;
                        if (attributeCollection.Contains("src"))
                        {
                            HtmlAttribute attribute = attributeCollection["src"];
                            srcValue = attribute.Value;
                        }
                        if (attributeCollection.Contains("alt"))
                        {
                            HtmlAttribute attribute = attributeCollection["alt"];
                            alt = attribute.Value;
                        }
                        m_resultDic.Add(alt, srcValue);
                        Console.WriteLine(alt + " " + srcValue);
                    }
                    InitResultDic(tempNode);
                }
            }
        }

        private static void DownloadImage(List<ItemData> iconDataList)
        {
            int count = iconDataList.Count;
            for (int i = 0; i < count; i++)
            {
                SaveImage(iconDataList[i].m_iconPath, iconDataList[i].m_id);
            }
        }

        private static void SaveImage(string url, int index)
        {
            WebClient mywebclient = new WebClient();
            string newfilename = index + ".jpg";
            string filepath = HtmlConst.kSaveImagePath + newfilename;
            mywebclient.DownloadFile(url, filepath);
        }

        private static void ExportExcel(string excelName, List<ItemData> iconDataList)
        {
            XlsDocument xls = new XlsDocument();
            xls.FileName = excelName;
            xls.SummaryInformation.Author = "daijialin";
            xls.SummaryInformation.Subject = "dotaHeroTexture";

            string sheetName = "Sheet0";
            org.in2bits.MyXls.Worksheet sheet = xls.Workbook.Worksheets.Add(sheetName);
            Cells cells = sheet.Cells;

            int rowNum = iconDataList.Count;
            int rowMin = 1;

            for (int i = 0; i < rowNum + 1; i++)
            {
                if (i == 0)
                {
                    cells.Add(1, 1, "id");
                    cells.Add(1, 2, "name");
                    cells.Add(1, 3, "headIcon");
                }
                else if (i == 1)
                {
                    cells.Add(2, 1, "");
                    cells.Add(2, 2, "名称");
                    cells.Add(2, 3, "头像");
                }
                else
                {
                    int currentRow = rowMin + i;
                    cells.Add(currentRow, 1, iconDataList[i - 1].m_id);
                    cells.Add(currentRow, 2, iconDataList[i - 1].m_heroName);
                    cells.Add(currentRow, 3, iconDataList[i - 1].m_iconPath);
                }

            }
            xls.Save();
        }

    }
}
