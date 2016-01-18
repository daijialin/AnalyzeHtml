using System;
using System.Collections.Generic;
using System.Net;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using org.in2bits.MyXls;

namespace Pengbo
{
    public class ParseHtml
    {
        private Dictionary<string, string> m_resultDic = new Dictionary<string, string>();
        private List<ItemData> m_itemDataList = new List<ItemData>();
        private string m_picPath = string.Empty;
        private string m_picName = string.Empty;
        public const string kSaveImagePath = @"D:\DotaResource\AccessoryIcons\";
        private const string kPicRoot = "bodyer";

        public ParseHtml() 
        {
        }

        public void Parse()
        {
            InitItemDataList();
            ExportExcel();
            ExportPic();
            Console.ReadLine();
        }

        private void InitItemDataList()
        {
            Range excelRange = ParseUtility.GetExcelRang(HtmlConst.kSourExcelPath);

            //Read Excel Info 
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
                string urlPath = soureUrlPath + HtmlConst.kChineaseTail;

                InitPicPathAndNameFromWeb(urlPath);

                if (!m_resultDic.ContainsKey(m_picName))
                {
                    m_resultDic.Add(m_picName, m_picPath);
                    AddItemDataToList(idStr, nameStr, qualityStr, positionStr, heroNameStr, m_picPath);
                }
                Console.WriteLine("Please Wait, Parsing Excel:  " + i + "" + m_picName);
            }
        }

        private  void InitPicPathAndNameFromWeb(string url)
        {
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(url);
            HtmlNode htmlNode = doc.GetElementbyId(kPicRoot);

            InitPicPathFromWeb(htmlNode);
            InitNameFromWeb(htmlNode);
        }

        private  void AddItemDataToList(string idStr, string nameStr, string qualityStr, string positionStr, string heroNameStr, string picUrl)
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

        private  void ExportPic()
        {
            int count = m_itemDataList.Count;
            for (int i = 0; i < count; i++)
            {
                DownLoadPic(m_itemDataList[i].m_iconPath, m_itemDataList[i].m_id);
                Console.WriteLine("Download is in progress, picName:" + "" + m_itemDataList[i].itemName);
            }
        }

        private  void DownLoadPic(string url, int index)
        {
            WebClient webClient = new WebClient();
            string picName = index + ".jpg";
            string filepath = HtmlConst.kSaveImagePath + picName;
            webClient.DownloadFile(url, filepath);
        }

        public static void InitCells(Cells cells, int rowNum, int rowMin, int line, object[] titles, object[] describes, List<ItemData> itemDataList)
        {
            for (int i = 0; i < rowNum + 2; i++)
            {
                if (i == 0)
                {
                    for (int j = 0; j < line; ++j)
                    {
                        cells.Add(1, j + 1, titles[j]);
                    }
                }
                else if (i == 1)
                {
                    for (int j = 0; j < line; ++j)
                    {
                        cells.Add(2, j + 1, describes[j]);
                    }
                }
                else
                {
                    int currentRow = rowMin + i;
                    cells.Add(currentRow, 1, itemDataList[i - 2].m_id);
                    cells.Add(currentRow, 2, itemDataList[i - 2].itemName);
                    cells.Add(currentRow, 3, itemDataList[i - 2].qualityStr);
                    cells.Add(currentRow, 4, itemDataList[i - 2].positionStr);
                    cells.Add(currentRow, 5, itemDataList[i - 2].m_heroName);
                    cells.Add(currentRow, 6, itemDataList[i - 2].m_iconPath);
                }
                Console.WriteLine("export Excel:" + "" + i);
            }
        }

        private void ExportExcel()
        {
            XlsDocument xls = ParseUtility.GetInitXlsDoc();
            org.in2bits.MyXls.Worksheet sheet0 = xls.Workbook.Worksheets[0];

            Cells cells = sheet0.Cells;

            int rowNum = m_itemDataList.Count;
            int rowMin = 1;
            int lineNum = 6;

            object[] titles = new object[6]
            {
                "id", "itemName",  "qualityStr", "positionStr", "heroName", "itemPath",
            };

            object[] describes = new object[6]
            {
                "", "饰品名称",  "饰品品质", "饰品位置", "饰品所属", "饰品路径",
            };


            InitCells(cells, rowNum, rowMin, 6, titles, describes, m_itemDataList);

            xls.Save(HtmlConst.kTargetExcelPath, true);
        }
    }

}
