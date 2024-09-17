namespace OneNoteTools
{
    public class Hyperlink
    {
        public enum HyperlinkTypeEnum
        {
            Unknown,
            File,
            Web,
            MailTo,
            OneNote
        }

        private string _name = string.Empty;
        private string _reference = string.Empty;
        private HyperlinkTypeEnum _linkType = HyperlinkTypeEnum.Unknown;

        public Hyperlink()
        { }

        public Hyperlink(string data)
        {

            // parse the text blob
            switch (data.Substring(0, 5).ToLower())
            {
                case "href=":
                    LoadHRef(data);
                    break;
                case "hyper":
                    LoadHyper(data);
                    break;
                default:
                    break;
            }

            // determine the protocol
            switch (Reference.GetInsideValue("", ":").ToLower())
            {
                case "onenote":
                    _linkType = HyperlinkTypeEnum.OneNote;
                    break;
                case "file":
                    _linkType = HyperlinkTypeEnum.File;
                    break;
                case "http":
                case "https":
                    _linkType = HyperlinkTypeEnum.Web;
                    break;
                case "mailto":
                    _linkType = HyperlinkTypeEnum.MailTo;
                    break;
                default:
                    break;
            }

        }

        private void LoadHRef(string value)
        {

            // remove the href=
            value = value.Substring(5);

            // remove the trailing close tag
            if (value.Substring(value.Length - 4) == "</a>")
                value = value.Substring(0, value.Length - 4);

            // if the reference is quoted
            if (value.Substring(0, 1) == @"""")
            {
                // pull out the first quoted value - this is the reference
                _reference = value.GetInsideValue(@"""", @"""", false);

                // grab the test of the value - this is the name
                _name = value.Substring(Reference.Length + 3);
            }
            else
            {
                // no quotes - split based on the last <

                int i = value.LastIndexOf("<");

                if (i > -1)
                {
                    _reference = value.Substring(0, i);
                    _name = value.Substring(i);
                }
            }

            // cleanup name if there are span's in it
            if (Name.IndexOf("<span") != -1)
            {
                if (Name.Substring(0, 6) == "<span ")
                {
                    int i = 0;
                    _name = Name.Substring(6);
                    _name = Name.Replace("</span>", "");
                    i = _name.IndexOf(">");
                    if (i != -1)
                    {
                        _name = Name.Substring(i + 1);
                    }
                }
            }
            if (_name.IndexOf("</span>") != -1)
                _name = _name.GetInsideValue(">", "</span>", false);

            // cleanup reference
            _reference = Reference.Replace("\r", "");
            _reference = Reference.Replace("\n", "");
            _name = Name.Trim();

        }

        private void LoadHyper(string value)
        {

            // for hyper link references, both the reference and the names are in quotes

            value = value.Substring(10);
            List<string> parts = value.GetInsideValues(@"""", @"""", false);
            if (parts.Count >= 1)
                _reference = parts[0];

            if (parts.Count >= 2)
                _name = parts[1];

        }

        #region properties

        public string Name { get { return _name; } }
        public string Reference { get { return _reference; } }
        public HyperlinkTypeEnum LinkType { get { return _linkType; } }
        public string PlainTextReference { get { return Reference.Decode(); } }

        public LinkInfo LinkInfo()
        {
            return new LinkInfo(_reference);
        }

        #endregion

    }
}
