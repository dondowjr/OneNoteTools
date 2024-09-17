using System.Text;

namespace OneNoteTools
{
    public class LinkInfo
    {
        public enum LinkTypeEnum
        {
            Unknown,
            File,
            Directory,
            Web,
            MailTo,
            Page,
            Section,
            Group,
            Notebook
        }

        #region local vars

        private string _linkValue = string.Empty;
        private string _fullPath = string.Empty;
        private string _sectionName = string.Empty;
        private string _pageName = string.Empty;
        private bool _externalLink = false;
        private bool _webLink = false;
        private LinkTypeEnum _linkType = LinkTypeEnum.Unknown;
        private string _externalPageID = string.Empty;
        private string _externalSectionID = string.Empty;
        private string _emailUserID = string.Empty;

        #endregion

        #region constuctors

        public LinkInfo(string linkValue)
        {

            _linkValue = linkValue;

            switch (_linkValue.GetInsideValue("", ":").ToLower())
            {
                case "onenote":
                    LoadOneNote(_linkValue);
                    break;
                case "file":
                    LoadFile(_linkValue);
                    break;
                case "http":
                case "https":
                    LoadWeb(_linkValue);
                    break;
                case "mailto":
                    LoadMailTo(_linkValue);
                    break;
                default:
                    break;
            }

        }

        private void LoadFile(string value)
        {

            // mostly just cleanup here
            _fullPath = value;

            if (_fullPath.Substring(0, 5) == "file:")
            {
                _fullPath = _fullPath.Substring(5);
                if (_fullPath.Substring(0, 3) == @"///")
                    _fullPath = _fullPath.Substring(3);

                _fullPath = _fullPath.Replace(@"/", @"\");
            }

            _linkType = LinkTypeEnum.File;

            // promote to directory if the file does not exist
            if (!File.Exists(_fullPath))
            {
                if (Directory.Exists(_fullPath))
                    _linkType = LinkTypeEnum.Directory;
            }

        }

        private void LoadWeb(string value)
        {
            _fullPath = value;
            _linkType = LinkTypeEnum.Web;
        }

        private void LoadMailTo(string value)
        {
            _fullPath = value;
            _linkType = LinkTypeEnum.MailTo;
        }

        #region load onenote

        private void LoadOneNote(string value)
        {
            // trim off the onenote: value
            if (value.Substring(0, 8).ToLower() == "onenote:")
                value = value.Substring(8);

            switch (value.Substring(0, 3))
            {
                case "///":
                    LoadExternalLocalLink(value);
                    break;
                case "htt":
                    LoadExternalWebLink(value);
                    break;
                default:
                    LoadInternalLink(value);
                    break;
            }

            // classify link based on what we parsed
            if (!string.IsNullOrEmpty(_externalPageID))
            {
                _linkType = LinkTypeEnum.Page;
            }
            else
            {
                if (!string.IsNullOrEmpty(_externalSectionID))
                {
                    _linkType = LinkTypeEnum.Section;
                }
            }

            if (_linkType == LinkTypeEnum.Unknown)
            {
                if (!string.IsNullOrEmpty(_fullPath))
                {
                    if (_fullPath.Length > 4)
                    {
                        if (_fullPath.Substring(_fullPath.Length - 4).ToLower() == ".one")
                            _linkType = LinkTypeEnum.Section;
                        else
                            _linkType = LinkTypeEnum.Group;
                    }
                    else
                    {
                        _linkType = LinkTypeEnum.Group;
                    }
                }
            }

        }
        private void LoadExternalLocalLink(string value)
        {


            // strip the file marker
            if (value.Substring(0, 3) == "///")
                value = value.Substring(3);

            // get the full path
            _fullPath = value.GetInsideValue("", "#");

            _fullPath = FullPath.Replace(@"/", @"\");

            _sectionName = _fullPath.Decode().GetInsideValue(@"\", ".one", false, true);

            // get the section name
            _pageName = value.GetInsideValue("#", "&").Decode();
            if (_pageName.IndexOf("-id={") != -1)
                _pageName = "";

            _externalPageID = value.GetInsideValue("page-id=", "&");
            _externalSectionID = value.GetInsideValue("section-id=", "&");

            _externalLink = true;

        }
        private void LoadExternalWebLink(string value)
        {

            //onenote:
            //        https://d.docs.live.net/8862a45af406527d/Documents/Work/Summits/MS%20Summit%202019.one#
            //        MS%20Summit%202019&
            //        section-id={03CB044E-A2B1-4B60-BA33-6801F4AD0002}&
            //        page-id={1D41153E-FDC7-4571-8045-12810B39C32E}&end

            // get the full path
            _fullPath = value.GetInsideValue("", "#");

            _sectionName = _fullPath.Decode().GetInsideValue(@"/", ".one", false, true);

            // get the section name
            _pageName = value.GetInsideValue("#", "&").Decode();
            if (_pageName.IndexOf("-id={") != -1)
                _pageName = "";

            _externalPageID = value.GetInsideValue("page-id=", "&");
            _externalSectionID = value.GetInsideValue("section-id=", "&");

            _webLink = true;
            _externalLink = true;

        }
        private void LoadInternalLink(string value)
        {
            //onenote: 
            //          Quick Notes.one#
            //          OneNote Basics&amp;
            //          section-id={056D3E22-03DE-4702-967C-15853C254E9B}&amp;
            //          page-id={CB7E1C93-D2B1-4E33-89EC-5BBE2784C5F2}&amp;
            //          base-path=//C:/Users/DRD/Documents/OneNote Notebooks/My Notebook

            //onenote:
            //          New Section Group/New Section 1.one#
            //          section-id={3B94651F-B088-4F39-889B-588701EA239C}&amp;
            //          page-id={D940A443-BD05-40CF-B544-164DAFFEC77E}&amp;
            //          base-path=//C:/Users/DRD/Documents/OneNote Notebooks/My Notebook

            string file = string.Empty;
            string path = string.Empty;

            // get the full path
            file = value.GetInsideValue("", "#");

            _sectionName = file.GetInsideValue("", ".one".Decode());
            _sectionName = _sectionName.Replace(@"/", @"\");

            // get the section name
            _pageName = value.GetInsideValue("#", "&").Decode();
            if (_pageName.IndexOf("-id={") != -1)
                _pageName = "";

            _externalPageID = value.GetInsideValue("page-id=", "&");
            _externalSectionID = value.GetInsideValue("section-id=", "&");

            path = value.GetInsideValue("base-path=", "");

            _fullPath = path + @"\" + file;
            _fullPath = _fullPath.Replace(@"/", @"\");

        }

        #endregion

        #endregion

        #region properties

        public string LinkValue { get { return _linkValue; } }
        public string FullPath { get { return _fullPath; } }
        public string FullPathPlainText
        { get { return _fullPath.Decode(); } }

        public string SectionName { get { return _sectionName; } }
        public string PageName { get { return _pageName; } }
        public bool ExternalLink { get { return _externalLink; } }
        public bool WebLink { get { return _webLink; } }
        public LinkTypeEnum LinkType { get { return _linkType; } }
        public string ExternalPageID { get { return _externalPageID; } }
        public string ExternalSectionID { get { return _externalSectionID; } }

        #endregion

        #region methods
        public override string ToString()
        {

            StringBuilder data = new StringBuilder();

            data.AppendLine("LinkValue: " + LinkValue);
            data.AppendLine("FullPath: " + FullPath);
            data.AppendLine("FullPathPlainText: " + FullPathPlainText);
            data.AppendLine("SectionName: " + SectionName);
            data.AppendLine("PageName: " + PageName);
            data.AppendLine("ExternalLink: " + ExternalLink);
            data.AppendLine("WebLink: " + WebLink);
            data.AppendLine("LinkType: " + LinkType);
            data.AppendLine("ExternalPageID: " + ExternalPageID);
            data.AppendLine("ExternalSectionID: " + ExternalSectionID);

            return data.ToString();

        }

        #endregion
    }
}
