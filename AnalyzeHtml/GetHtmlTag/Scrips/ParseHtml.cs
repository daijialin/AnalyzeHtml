using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using org.in2bits.MyXls;

namespace GetHtmlTag
{
    public class ParseHtml
    {
        private Dictionary<string, string> m_resultDic = new Dictionary<string, string>();
        private List<ItemData> m_itemDataList = new List<ItemData>();
        private string path = @"E:\\pengbo\\github\\AnalyzeHtml\\AnalyzeHtml\\GetHtmlTag\\bin\\Debug\\OrnamentsConfig.xls";
        private string m_srcValue = string.Empty;
        private string m_alt = string.Empty;

        public ParseHtml() 
        {
            
        }

        public  void ReadXml()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wbook = app.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)wbook.Worksheets[1];
            string info = ((Range)worksheet.Cells[2, 2]).Text.ToString();
            Range range = worksheet.Cells;

            for (int i = 0; i < range.Count; ++i)
            {
                int index = i + 2;
                string cellPicPath = range[index, 6].Text.ToString();
                Console.WriteLine(cellPicPath);
                DoHtml(cellPicPath);
                break;
                if (string.IsNullOrEmpty(cellPicPath))
                {
                    break;
                }
            }
            Console.ReadLine();
        }

        private const string kPicRoot = "bodyer";

        private  void DoHtml(string url)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(url);
            HtmlNode htmlNode = doc.GetElementbyId(kPicRoot);

            InitPicPath(htmlNode);
            InitName(htmlNode);

            m_resultDic.Add(m_alt, m_srcValue);

            InitIconDataList();
            ExportExcel("accessoryExcel", m_itemDataList);
            DownloadImage(m_itemDataList);
            //Console.WriteLine("ok");
        }

        private  void InitIconDataList()
        {  
            int index = 0;
            foreach (KeyValuePair<string, string> item in m_resultDic)
            {
                ItemData itemData = new ItemData();
                itemData.m_id = index;
                itemData.m_heroName = item.Key;
                itemData.m_iconPath = item.Value;
                index++;
                m_itemDataList.Add(itemData);
            }
        }




        private  void InitPicPath(HtmlNode nodeA)
        {
            if (nodeA.HasChildNodes)
            {
                for (int i = 0; i < nodeA.ChildNodes.Count; i++)
                {
                    HtmlNode tempNode = nodeA.ChildNodes[i];
                    HtmlNode childNode = ParseUtility.GetNode(nodeA, "div", "class", "market-box-open");
                    if (childNode != null)
                    {
                        for (int j = 0; j < childNode.ChildNodes.Count; ++j)
                        {
                            
                            HtmlNode tempChildNode = childNode.ChildNodes[j];
                            HtmlAttribute htmlAttribute = ParseUtility.GetAttribute(tempChildNode, "img", "src");
                            if (htmlAttribute != null)
                            {
                                m_srcValue = htmlAttribute.Value;
                            }
                        }
                    }

                    InitPicPath(tempNode);               
                }
            }
        }

        private void InitName(HtmlNode soureNode)
        {
            if (soureNode.HasChildNodes)
            {
                for (int i = 0; i < soureNode.ChildNodes.Count; i++)
                {
                    HtmlNode node0 = soureNode.ChildNodes[i];

                    HtmlNode childNode = ParseUtility.GetNode(node0, "h4", "class", "market-article-tt");

                    if (childNode != null)
                    {
                        m_alt = childNode.InnerText;
                    }
                    InitName(node0);

                }
            }
        }

        private  void DownloadImage(List<ItemData> iconDataList)
        {
            int count = iconDataList.Count;
            for (int i = 0; i < count; i++)
            {
                SaveImage(iconDataList[i].m_iconPath, iconDataList[i].m_id);
            }
        }

        public const string kSaveImagePath = @"D:\DotaResource\AccessoryIcons\";
        private  void SaveImage(string url, int index)
        {
            WebClient mywebclient = new WebClient();
            string newfilename = index + ".jpg";
            string filepath = kSaveImagePath + newfilename;
            mywebclient.DownloadFile(url, filepath);
        }

        private  void ExportExcel(string excelName, List<ItemData> iconDataList)
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

            for (int i = 0; i < rowNum + 2; i++)
            {
                if (i == 0)
                {
                    cells.Add(1, 1, "id");
                    cells.Add(1, 2, "itemName");
                    cells.Add(1, 3, "itemPath");
                }
                else if (i == 1)
                {
                    cells.Add(2, 1, "");
                    cells.Add(2, 2, "饰品");
                    cells.Add(2, 3, "饰品路径");
                }
                else
                {
                    int currentRow = rowMin + i;
                    cells.Add(currentRow, 1, iconDataList[i - 2].m_id);
                    cells.Add(currentRow, 2, iconDataList[i - 2].m_heroName);
                    cells.Add(currentRow, 3, iconDataList[i - 2].m_iconPath);
                }

            }
            xls.Save();
        }
    }

}
