using System.Xml;

namespace OneNoteTools
{
    /// <summary>
    /// Lightweight object that contains primary data about an OneNote page
    /// </summary>
    public class PageInfo
    {

        #region local vars

        private string _name;
        private string _path;
        private string _id;
        private DateTime _dateModified;
        private DateTime _dateCreated;

        #endregion

        #region constructor

        public PageInfo(XmlNode xml)
        {

            LoadByXML(xml);

        }

        private void LoadByXML(XmlNode xml)
        {

            foreach (XmlAttribute item in xml.Attributes)
            {
                switch (item.Name.ToLower())
                {
                    case "name":
                        _name = item.Value;
                        break;
                    case "id":
                        _id = item.Value;
                        break;
                    case "path":
                        _path = item.Value;
                        break;
                    case "lastmodifiedtime":
                        DateTime.TryParse(item.Value, out _dateModified);
                        break;
                    case "datetime":
                        DateTime.TryParse(item.Value, out _dateCreated);
                        break;
                    default:
                        break;
                }
            }

            if (xml.ParentNode != null)
            {
                if (xml.ParentNode.Attributes != null)
                {
                    if (xml.ParentNode.Attributes.Count != 0)
                    {
                        foreach (XmlAttribute item in xml.ParentNode.Attributes)
                        {
                            if (item.Name == "path")
                                _path = item.Value;
                        }
                    }
                }
            }

        }

        #endregion

        #region properties

        public string ID
        { get { return _id; } }

        public string Name
        { get { return _name; } }

        public string Path
        { get { return _path; } }

        public DateTime DateModified
        { get { return _dateModified; } }

        public DateTime DateCreated
        { get { return _dateCreated; } }

        #endregion

    }
}
