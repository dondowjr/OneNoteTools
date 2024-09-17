using System.Xml;

namespace OneNoteTools
{
    public class Notebook
    {
        private string _id;
        private string _name;
        private string _path;
        private DateTime _dateModified;

        public Notebook(XmlNode xml)
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
                    default:
                        break;
                }
            }

        }

        public string ID
        { get { return _id; } }

        public string Name
        { get { return _name; } }

        public string Path
        { get { return _path; } }

        public DateTime LastEdit
        { get { return _dateModified; } }

    }
}
