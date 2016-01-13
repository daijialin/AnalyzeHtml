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
        private string m_picPath = string.Empty;
        private string m_picName = string.Empty;
        public const string kSaveImagePath = @"D:\DotaResource\AccessoryIcons\";
        private const string kPicRoot = "bodyer";

        public ParseHtml() 
        {
        }

        public  void ReadXml()
        {
            Range excelRange = ParseUtility.GetExcelRang(path);

            for (int i = 0; i < excelRange.Count; ++i)
            {
                int index = i + 2;
                string idStr = excelRange[index, 1].Text.ToString();
                string nameStr = excelRange[index, 2].Text.ToString();
                string qualityStr = excelRange[index, 3].Text.ToString();
                string positionStr = excelRange[index, 4].Text.ToString();
                string heroNameStr = excelRange[index, 5].Text.ToString();

                string soureUrlPath = excelRange[index, 6].Text.ToString();

                if (string.IsNullOrEmpty(soureUrlPath))
                {
                    break;
                }
                string urlPath = soureUrlPath + "?lang=zh_cn";

                InitPicAndNameFromWeb(urlPath);

                if (!m_resultDic.ContainsKey(m_picName))
                {
                    m_resultDic.Add(m_picName, m_picPath);
                    InitIconDataList(idStr, nameStr, qualityStr, positionStr, heroNameStr, m_picPath);
                }
                Console.WriteLine("ParseWeb:  " + i + "" + m_picName);
            }

            ExportExcel(m_itemDataList, "AccessoryExcel");
            DownloadImage(m_itemDataList);
            Console.ReadLine();
        }

        private  void InitPicAndNameFromWeb(string url)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(url);
            HtmlNode htmlNode = doc.GetElementbyId(kPicRoot);

            InitPicPathFromWeb(htmlNode);
            InitNameFromWeb(htmlNode);
        }

        private  void InitIconDataList(string idStr, string nameStr, string qualityStr, string positionStr, string heroNameStr, string picUrl)
        {
            ItemData itemData = new ItemData();
            itemData.m_id = 100000 + int.Parse(idStr);
            itemData.itemName = nameStr;
            itemData.qualityStr = qualityStr;
            itemData.positionStr = positionStr;
            itemData.m_heroName = heroNameStr;
            itemData.m_iconPath = picUrl;

            //index++;
            m_itemDataList.Add(itemData);
        }
        private  void InitPicPathFromWeb(HtmlNode nodeA)
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
                                m_picPath = htmlAttribute.Value;
                            }
                        }
                    }

                    InitPicPathFromWeb(tempNode);               
                }
            }
        }

        private void InitNameFromWeb(HtmlNode soureNode)
        {
            if (soureNode.HasChildNodes)
            {
                for (int i = 0; i < soureNode.ChildNodes.Count; i++)
                {
                    HtmlNode node0 = soureNode.ChildNodes[i];

                    HtmlNode childNode = ParseUtility.GetNode(node0, "h4", "class", "market-article-tt");

                    if (childNode != null)
                    {
                        m_picName = childNode.InnerText;
                    }
                    InitNameFromWeb(node0);

                }
            }
        }

        private  void DownloadImage(List<ItemData> iconDataList)
        {
            int count = iconDataList.Count;
            for (int i = 0; i < count; i++)
            {
                SaveImage(iconDataList[i].m_iconPath, iconDataList[i].m_id);
                Console.WriteLine("downLoadPic:" + "" + iconDataList[i].itemName);
            }
        }

        private  void SaveImage(string url, int index)
        {
            WebClient mywebclient = new WebClient();
            string newfilename = index + ".jpg";
            string filepath = kSaveImagePath + newfilename;
            mywebclient.DownloadFile(url, filepath);
        }

        private void ExportExcel(List<ItemData> iconDataList, string excelName = "newExcel", string author = "admin", string subject = "admin", string sheetName = "Sheet0")
        {
            XlsDocument xls = new XlsDocument();
            xls.FileName = excelName;
            xls.SummaryInformation.Author = author;
            xls.SummaryInformation.Subject = subject;

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
                    cells.Add(1, 3, "qualityStr");
                    cells.Add(1, 4, "positionStr");
                    cells.Add(1, 5, "heroName");
                    cells.Add(1, 6, "itemPath");
                }
                else if (i == 1)
                {
                    cells.Add(2, 1, "");
                    cells.Add(2, 2, "饰品名称");
                    cells.Add(2, 3, "饰品品质");
                    cells.Add(2, 4, "饰品位置");
                    cells.Add(2, 5, "饰品所属");
                    cells.Add(2, 6, "饰品URL路径");
                }
                else
                {
                    int currentRow = rowMin + i;
                    cells.Add(currentRow, 1, iconDataList[i - 2].m_id);
                    cells.Add(currentRow, 2, iconDataList[i - 2].itemName);
                    cells.Add(currentRow, 3, iconDataList[i - 2].qualityStr);
                    cells.Add(currentRow, 4, iconDataList[i - 2].positionStr);
                    cells.Add(currentRow, 5, iconDataList[i - 2].m_heroName);
                    cells.Add(currentRow, 6, iconDataList[i - 2].m_iconPath);
                }
                Console.WriteLine("export Excel:" + "" + i);

            }
            xls.Save();
        }
    }

}
