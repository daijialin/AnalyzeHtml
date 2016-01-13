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
        }
    }
}