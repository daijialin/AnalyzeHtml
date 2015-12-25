using HtmlAgilityPack;

namespace GetHtmlTag
{
    public static class ParseUtility
    {
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
