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

namespace GetHtmlTag
{
    class Program
    {
        //private static string m_url = "http://db.dota2.uuu9.com/simulator/index";
        private static string m_url = "http://dota2.vpgame.com/market.html?lang=zh_cn";
        private const string UTF = "utf-8";
        private const string GB = "gb2312";
        private static List<TextureData> m_TextureDataList;
        private static Dictionary<string, string> resultDiction = new Dictionary<string, string>();

        //static void Main(string[] args)
        //{
        //    resultDiction = new Dictionary<string, string>();
        //    DoHtmlDota2();
        //}

        private static void GetHtmlNode(HtmlNode mainNode)
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
                        resultDiction.Add(alt, srcValue);
                        Console.WriteLine(alt + " " + srcValue);
                    }
                    GetHtmlNode(tempNode);
                }
            }
        }

        private static void DoHtmlDota2()
        {
            m_TextureDataList = new List<TextureData>();
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(m_url);
            HtmlNode mainNode = doc.GetElementbyId("market-sorts-hero");
            GetHtmlNode(mainNode);

            List<TextureData> textureDataList = new List<TextureData>();
            int index=0;
            foreach (KeyValuePair<string, string> item in resultDiction)
            {
                TextureData td = new TextureData();
                td.id = index;
                td.heroName = item.Key;
                td.icoPath = item.Value;
                index++;
                textureDataList.Add(td);
            }
            ExportExcel("smallHeroIcons",textureDataList);
            DownloadImage(textureDataList);
            Console.WriteLine("ok");
        }


        private static void DoHtmlDota1()
        {
            m_TextureDataList = new List<TextureData>();
            HtmlWeb web = new HtmlWeb();
            web.OverrideEncoding = Encoding.GetEncoding("gb2312");
            web.AutoDetectEncoding = false;
            HtmlDocument doc = web.Load(m_url);
            doc.OptionDefaultStreamEncoding = Encoding.GetEncoding("gb2312");
            int index = 0;
            HtmlNode mainNode = doc.GetElementbyId("herolist1");
            Console.OutputEncoding = Encoding.UTF8;
            HtmlNodeCollection childNodes = mainNode.ChildNodes;
            for (int i = 0; i < childNodes.Count; i++)
            {
                HtmlNode tempNode = childNodes[i];
                if (tempNode.Name.Equals("a"))
                {
                    HtmlDocument document = new HtmlDocument();
                    document.LoadHtml(tempNode.InnerHtml);
                    var type = typeof(HtmlDocument).Assembly.GetType("HtmlAgilityPack.HtmlDocument");
                    FieldInfo infoNames = type.GetField("_currentattribute", BindingFlags.NonPublic | BindingFlags.Instance);
                    HtmlAttribute value = (HtmlAttribute)infoNames.GetValue(document);
                    HtmlNode nextNode = value.OwnerNode.NextSibling;
                    Console.OutputEncoding = Encoding.GetEncoding("gb2312");
                    TextureData data = new TextureData();
                    data.id = index;
                    data.icoPath = value.Value;
                    data.heroName = nextNode.InnerHtml;
                    m_TextureDataList.Add(data);
                    index++;
                }
            }
            //ExportExcel("test", m_TextureDataList);
            //DownloadImage(m_TextureDataList);
            Console.WriteLine("ok");
        }


        private static void DownloadImage(List<TextureData> textureDataList)
        {
            int count = textureDataList.Count;
            for (int i = 0; i < count; i++)
            {
                SaveImage(textureDataList[i].icoPath, textureDataList[i].id);
            }
        }

        private static void SaveImage(string url, int index)
        {
            WebClient mywebclient = new WebClient();
            string newfilename = index + ".jpg";
            string filepath = @"D:\DotaResource\SmallHeroIcons\" + newfilename;
            mywebclient.DownloadFile(url, filepath);
        }

        private static void ExportExcel(string name, List<TextureData> textureDataList)
        {
            XlsDocument xls = new XlsDocument();
            xls.FileName = name;
            xls.SummaryInformation.Author = "daijialin";
            xls.SummaryInformation.Subject = "dotaHeroTexture";

            string sheetName = "Sheet0";
            Worksheet sheet = xls.Workbook.Worksheets.Add(sheetName);
            Cells cells = sheet.Cells;

            int rowNum = textureDataList.Count;
            int rowMin = 1;

            for (int i = 0; i < rowNum + 1; i++)
            {
                if (i == 0)
                {
                    cells.Add(1, 1, "m_id");
                    cells.Add(1, 2, "name");
                    cells.Add(1, 3, "path");
                }
                else
                {
                    int currentRow = rowMin + i;
                    cells.Add(currentRow, 1, textureDataList[i - 1].id);
                    cells.Add(currentRow, 2, textureDataList[i - 1].heroName);
                    cells.Add(currentRow, 3, textureDataList[i - 1].icoPath);
                }

            }
            xls.Save();
        }

    }

    public class TextureData
    {
        public int id;
        public string icoPath;
        public string heroName;
    }
}
