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
        private List<IconData> iconDataList = new List<IconData>();
        private string path = @"E:\\pengbo\\github\\AnalyzeHtml\\AnalyzeHtml\\GetHtmlTag\\bin\\Debug\\OrnamentsConfig.xls";

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

            m_resultDic.Add(alt, srcValue);

            InitIconDataList();
            ExportExcel(HtmlConst.kExportExcelName, iconDataList);
            DownloadImage(iconDataList);
            //Console.WriteLine("ok");
        }

        private  void InitIconDataList()
        {
       
            int index = 0;
            foreach (KeyValuePair<string, string> item in m_resultDic)
            {
                IconData td = new IconData();
                td.m_id = index;
                td.m_heroName = item.Key;
                td.m_iconPath = item.Value;
                index++;
                iconDataList.Add(td);
            }
        }

        private string srcValue = string.Empty;
        private string alt = string.Empty;


        private  void InitPicPath(HtmlNode nodeA)
        {
            if (nodeA.HasChildNodes)
            {

                for (int i = 0; i < nodeA.ChildNodes.Count; i++)
                {
                    HtmlNode tempNode = nodeA.ChildNodes[i];
                    HtmlNode childNode = GetNode(nodeA, "div", "class", "market-box-open");
                    if (childNode != null)
                    {
                        for (int j = 0; j < childNode.ChildNodes.Count; ++j)
                        {
                            
                            HtmlNode tempChildNode = childNode.ChildNodes[j];
                            HtmlAttribute htmlAttribute = GetAttribute(tempChildNode, "img", "src");
                            if (htmlAttribute != null)
                            {
                                srcValue = htmlAttribute.Value;
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

                    HtmlNode childNode = GetNode(node0, "h4","class", "market-article-tt");

                    if (childNode != null)
                    {
                        alt = childNode.InnerText;
                    }
                    InitName(node0);

                }
            }
        }

        private HtmlAttribute GetAttribute(HtmlNode sourceNode, string head, string tail)
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

        private HtmlNode GetNode(HtmlNode sourceNode, string head, string tail, string tailName)
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

        private  void DownloadImage(List<IconData> iconDataList)
        {
            int count = iconDataList.Count;
            for (int i = 0; i < count; i++)
            {
                SaveImage(iconDataList[i].m_iconPath, iconDataList[i].m_id);
            }
        }

        private  void SaveImage(string url, int index)
        {
            WebClient mywebclient = new WebClient();
            string newfilename = index + ".jpg";
            string filepath = HtmlConst.kSaveImagePath + newfilename;
            mywebclient.DownloadFile(url, filepath);
        }

        private  void ExportExcel(string excelName, List<IconData> iconDataList)
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
                    cells.Add(1, 2, "tailName");
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
