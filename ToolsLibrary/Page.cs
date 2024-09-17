using System.Xml;

namespace OneNoteTools
{
    /// <summary>
    /// Information about a OneNote page
    /// </summary>
    public class Page
    {

        #region local vars

        private string _id;
        private string _name;
        private DateTime _dateCreated = DateTime.MinValue;
        private DateTime _dateModified = DateTime.MinValue;
        private string _createAuthor;
        private string _modifiedAuthor;
        private List<Hyperlink> _hyperLinks = new List<Hyperlink>();

        #endregion

        #region constructors

        public Page()
        { }

        public Page(XmlNode xml)
        {

            LoadByXML(xml);

        }

        public void LoadByXML(XmlNode xml)
        {

            _id = string.Empty;
            _name = string.Empty;
            _dateCreated = DateTime.MinValue;
            _dateModified = DateTime.MinValue;
            _createAuthor = string.Empty;
            _modifiedAuthor = string.Empty;
            _hyperLinks.Clear();

            GetGeneralInfo(xml);
            GetAuthorAndDates(xml);
            GetLinks(xml);

        }

        private void GetGeneralInfo(XmlNode xml)
        {

            // get general page info
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
                    //case "path":
                    //    _path = item.Value;
                    //    break;
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

        }

        private void GetAuthorAndDates(XmlNode xml)
        {

            // get page authors and dates (which is a child of the page object)
            if (xml.Name.ToLower() == "one:title")
            {
                foreach (XmlNode item in xml.ChildNodes)
                {
                    if (item.Name == "one:OE")
                    {
                        foreach (XmlAttribute attr in item.Attributes)
                        {
                            switch (attr.Name.ToLower())
                            {
                                case "author":
                                    _createAuthor = attr.Value;
                                    break;
                                case "lastmodifiedby":
                                    _modifiedAuthor = attr.Value;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
            else
            {
                foreach (XmlNode child in xml.ChildNodes)
                    GetAuthorAndDates(child);
            }

        }

        private void GetLinks(XmlNode xml)
        {
            // todo: finish this method?
        }

        #endregion

        #region Properties

        public string ID
        { get { return _id; } }

        //public string PageContent
        //{
        //    get { return _pageContent; }
        //    set { _pageContent = value; }
        //}

        public string Name
        { get { return _name; } }

        //public string Path
        //{ get { return _path; } }

        public DateTime DateCreated
        { get { return _dateCreated; } }

        public DateTime DateModified
        { get { return _dateModified; } }

        public string CreateAuthor
        { get { return _createAuthor; } }

        public string ModifyAuthor
        { get { return _modifiedAuthor; } }

        public List<Hyperlink> HyperLinks
        { get { return _hyperLinks; } }

        #endregion

    }
}
