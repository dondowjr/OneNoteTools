using System.Xml;

namespace OneNoteTools
{
    /// <summary>
    /// This class contains helper routines for parsing OneNote XML data.
    /// </summary>
    public static class XMLHelper
    {

        public static void GetNameIDPairs(string xmlData, string xmlNodeNames, ref SortedList<string, string> items)
        {

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);
            List<string> NodeNames = xmlNodeNames.Split(";".ToCharArray()).ToList<string>();

            foreach (string name in NodeNames)
            {
                foreach (XmlNode item in xml.GetElementsByTagName(name))
                    items.Add(item.Attributes["name"].Value, item.Attributes["ID"].Value);
            }
        }

        public static void GetNameIDPairs(string xmlData, string xmlNodeNames, ref List<NameValue> items)
        {

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);
            List<string> NodeNames = xmlNodeNames.Split(";".ToCharArray()).ToList<string>();

            foreach (string name in NodeNames)
            {
                foreach (XmlNode item in xml.GetElementsByTagName(name))
                    items.Add(new NameValue(item.Attributes["name"].Value, item.Attributes["ID"].Value));
            }
        }

        public static void GetNotebooks(string xmlData, ref List<Notebook> items)
        {

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);

            foreach (XmlNode item in xml.GetElementsByTagName("one:Notebook"))
                items.Add(new Notebook(item));

        }

        public static void GetPagesInfo(string xmlData, ref List<PageInfo> items)
        {

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);

            foreach (XmlNode item in xml.GetElementsByTagName("one:Page"))
                items.Add(new PageInfo((XmlNode)item));

        }

        public static void GetPage(string xmlData, ref Page page)
        {
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(xmlData);

            if (xml.HasChildNodes)
            {
                foreach (XmlNode item in xml.ChildNodes)
                    GetPage(item, ref page);
            }

            GetHyperLinks(xmlData, ref page);

        }

        public static void GetPage(XmlNode xml, ref Page page)
        {
            if (xml.Name == "one:Page")
                page.LoadByXML(xml);
            else
            {
                if (xml.HasChildNodes)
                {
                    foreach (XmlNode item in xml.ChildNodes)
                        GetPage(item, ref page);
                }
            }
        }

        public static void GetHyperLinks(string xmlData, ref Page page)
        {

            if (xmlData.IndexOf("href=") != -1)
            {
                foreach (string value in xmlData.GetInsideValues("href=", @"/a>", true))
                    page.HyperLinks.Add(new Hyperlink(value));
            }

            if (xmlData.IndexOf("hyperlink=") != -1)
            {
                foreach (string value in xmlData.GetInsideValues("hyperlink=", ">", true))
                    page.HyperLinks.Add(new Hyperlink(value));
            }

        }

        public static void GetHyperLinksFromPageData(string xmlData, ref List<NameValue> list)
        {
            Hyperlink item = null;
            if (xmlData.IndexOf("href=") != -1)
            {
                foreach (string value in xmlData.GetInsideValues("href=", @"/a>", true))
                {
                    item = new Hyperlink(value);
                    list.Add(new NameValue(item.Name, item.Reference));
                    item = null;
                }
            }

            if (xmlData.IndexOf("hyperlink=") != -1)
            {
                foreach (string value in xmlData.GetInsideValues("hyperlink=", ">", true))
                {
                    item = new Hyperlink(value);
                    list.Add(new NameValue(item.Name, item.Reference));
                    item = null;
                }
            }
        }

    }

}
