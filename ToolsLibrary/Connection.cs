using Microsoft.Office.Interop.OneNote;
using System.Runtime.InteropServices;

namespace OneNoteTools
{/// <summary>
/// The connection object is the top-most or start object for all other classes. 
/// /// </summary>
    public class Connection : IDisposable
    {
        public enum OneNoteObjType
        {
            Unknown,
            Notebook,
            SectionGroup,
            Section,
            Page
        }

        #region - local vars -

        private Application2 _oneNote;
        private bool _disposed = false;

        #endregion

        #region - contructor - 

        public Connection()
        {
            _oneNote = new Application2();
        }

        #endregion

        #region - IDisposable support - 

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
                Cleanup();
        }

        ~Connection()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Cleanup()
        {
            if (!_disposed)
            {
                while (Marshal.ReleaseComObject(_oneNote) > 0) ;
                _oneNote = null;
                _disposed = true;
            }
        }

        #endregion

        #region - current items -
        // methods to get objects that are currently display or running in the onemote client - 

        /// <summary>
        /// Gets the Page ID for the currently displayed page.
        /// </summary>
        public string GetCurrentPageID()
        {
            if (_oneNote.Windows.Count != 0)
                return _oneNote.Windows.CurrentWindow.CurrentPageId;
            else
                return string.Empty;
        }

        /// <summary>
        /// Gets the Page object for the currently displayed page.
        /// </summary>
        public Page GetCurrentPage()
        {
            Page item = null;
            if (_oneNote.Windows.Count != 0)
            {
                string itemId = _oneNote.Windows.CurrentWindow.CurrentPageId;
                string data = null;
                _oneNote.GetPageContent(itemId, out data);
                item = new Page();
                XMLHelper.GetPage(data, ref item);
            }
            return item;
        }

        /// <summary>
        /// Gets the PAgeInfo object for thge currently displayed page.
        /// </summary>
        /// <returns></returns>
        public PageInfo GetCurrentPageInfo()
        {
            PageInfo item = null;

            if (_oneNote.Windows.Count != 0)
            {
                string itemId = _oneNote.Windows.CurrentWindow.CurrentPageId;

                item = GetPageInfo(itemId);

            }

            return item;
        }

        /// <summary>
        /// Gets the notebook ID for the currently displayed page.
        /// </summary>
        public string GetCurrentNotebookID()
        {
            if (_oneNote.Windows.Count != 0)
                return _oneNote.Windows.CurrentWindow.CurrentNotebookId;
            else
                return string.Empty;

        }

        /// <summary>
        /// Gets the notebook object for the currently displayed page.
        /// </summary>
        public Notebook GetCurrentNotebook()
        {

            Notebook item = null;

            if (_oneNote.Windows.Count != 0)
                return GetNotebook(_oneNote.Windows.CurrentWindow.CurrentNotebookId);

            return item;

        }

        /// <summary>
        /// Gets the sectoin ID for the currently displayed page.
        /// </summary>
        /// <returns></returns>
        public string GetCurrentSectionID()
        {
            if (_oneNote.Windows.Count != 0)
                return _oneNote.Windows.CurrentWindow.CurrentSectionId;
            else
                return string.Empty;
        }

        /// <summary>
        /// Gets the sectoin group ID for the currently displayed page.
        /// </summary>
        /// <returns></returns>
        public string GetCurrentSectionGroupID()
        {
            if (_oneNote.Windows.Count != 0)
                return _oneNote.Windows.CurrentWindow.CurrentSectionGroupId;
            else
                return string.Empty;
        }

        #endregion

        #region - notebooks -
        // methods to get notebook references from notenote

        /// <summary>
        /// Returns a notebook object for the given notebook ID.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public Notebook GetNotebook(string id)
        {
            Notebook item = null;

            // get the xml blob
            string data = null;
            _oneNote.GetHierarchy(id, HierarchyScope.hsNotebooks, out data);

            // parse xml
            List<Notebook> items = new List<Notebook>();
            XMLHelper.GetNotebooks(data, ref items);

            if (items.Count == 1)
                item = items[0];

            return item;
        }

        /// <summary>
        /// Gets a notebook list collection for the currently subscribed notebooks 
        /// </summary>
        /// <returns></returns>
        public List<Notebook> GetNotebooks()
        {

            List<Notebook> items = new List<Notebook>();

            string data = null;

            _oneNote.GetHierarchy(null, HierarchyScope.hsNotebooks, out data);

            XMLHelper.GetNotebooks(data, ref items);

            return items;

        }

        /// <summary>
        /// Returns the name and ID pairs for the currently subscribed notebooks 
        /// </summary>
        /// <returns></returns>
        public List<NameValue> GetNotebookNameAndIDs()
        {

            List<NameValue> items = new List<NameValue>();

            string data = null;

            _oneNote.GetHierarchy(null, HierarchyScope.hsNotebooks, out data);

            XMLHelper.GetNameIDPairs(data, "one:Notebook", ref items);

            return items;

        }

        #endregion

        #region - pageinfo -
        // methods to get pageinfo objects

        public PageInfo GetPageInfo(string id)
        {

            PageInfo item = null;
            List<PageInfo> items = new List<PageInfo>();

            string data = null;

            _oneNote.GetHierarchy(id, HierarchyScope.hsPages, out data);

            XMLHelper.GetPagesInfo(data, ref items);
            if (items.Count == 1)
                item = items[0];

            return item;

        }

        public List<PageInfo> GetPagesInfo()
        {
            return GetPagesInfo(string.Empty);
        }

        public List<PageInfo> GetPagesInfo(Notebook notebook)
        {

            List<PageInfo> items = new List<PageInfo>();

            string data = null;
            if (notebook == null)
                _oneNote.GetHierarchy(null, HierarchyScope.hsPages, out data);
            else
                _oneNote.GetHierarchy(notebook.ID, HierarchyScope.hsPages, out data);

            XMLHelper.GetPagesInfo(data, ref items);

            return items;

        }

        /// <summary>
        /// Gets a list of pages starting with and below the given parentID. If the parentID is null or blank, all pages are returned for all notebooks.
        /// 
        /// </summary>
        /// <param name="parentID"></param>
        /// <returns></returns>
        public List<PageInfo> GetPagesInfo(string parentID)
        {

            List<PageInfo> items = new List<PageInfo>();

            string data = null;
            if (string.IsNullOrEmpty(parentID))
                _oneNote.GetHierarchy(null, HierarchyScope.hsPages, out data);
            else
                _oneNote.GetHierarchy(parentID, HierarchyScope.hsPages, out data);

            XMLHelper.GetPagesInfo(data, ref items);

            return items;

        }

        /// <summary>
        /// Gets a list of all page names and ID.
        /// </summary>
        /// <returns></returns>
        public List<NameValue> GetPageNameAndIDs()
        {
            return GetPageNameAndIDs(null);
        }

        /// Gets a list of page names and ID that are found at or below the parentID.
        public List<NameValue> GetPageNameAndIDs(string parentID)
        {

            List<NameValue> items = new List<NameValue>();

            string data = null;

            if (string.IsNullOrEmpty(parentID))
                _oneNote.GetHierarchy(null, HierarchyScope.hsPages, out data);
            else
                _oneNote.GetHierarchy(parentID, HierarchyScope.hsPages, out data);

            XMLHelper.GetNameIDPairs(data, "one:Page", ref items);

            return items;
        }

        #endregion

        #region - page - 
        // methods to get page objects 

        public Page GetPage(string id)
        {
            string data = null;

            Page item = null;

            _oneNote.GetPageContent(id, out data, Microsoft.Office.Interop.OneNote.PageInfo.piAll);
            item = new Page();
            XMLHelper.GetPage(data, ref item);

            return item;
        }

        public Page GetPage(PageInfo item)
        {
            return GetPage(item.ID);
        }

        /// <summary>
        /// Gets a list of hyperlink names and hyperlink paths found on a page given the pages internal ID.  
        /// </summary>
        /// <param name="id">Internal ID</param>
        /// <returns></returns>
        public List<NameValue> GetHyperlinksOnPage(string id)
        {

            List<NameValue> items = new List<NameValue>();

            string data = null;

            _oneNote.GetPageContent(id, out data, Microsoft.Office.Interop.OneNote.PageInfo.piAll);

            XMLHelper.GetHyperLinksFromPageData(data, ref items);

            return items;

        }

        /// <summary>
        /// Updates the page data.
        /// </summary>
        /// <param name="id">ID of the page being updated. Used to validate the correct page os being updated.</param>
        /// <param name="oldPageData">Original page data. Used to validate that the page has not changed since the page data was last fetched.</param>
        /// <param name="newPageData">New page data.</param>
        public void UpdatePageData(string id, string oldPageData, string newPageData)
        {

            string data = GetPageData(id);

            if (data != oldPageData)
                throw new Exception("Data has changed.  Cannot update page.");

            _oneNote.UpdatePageContent(newPageData);

        }

        /// <summary>
        /// Gets the page data XML given the internal page ID.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetPageData(string id)
        {

            string data;
            _oneNote.GetPageContent(id, out data, Microsoft.Office.Interop.OneNote.PageInfo.piAll);
            return data;

        }

        #endregion

        #region - section -
        // methods to get onenote section information 


        /// <summary>
        /// Gets the section ID based on a path.
        /// </summary>
        /// <param name="plainTextPath">Path to the section, without encoding.</param>
        /// <returns></returns>
        public string GetSectionID(string plainTextPath)
        {

            string notebookID = string.Empty;
            string data = null;
            string sectionID = string.Empty;

            plainTextPath = plainTextPath.Replace(@"/", @"\");

            if (plainTextPath.Right(4) == ".one")
                plainTextPath = plainTextPath.CropRight(4);

            List<string> parts = plainTextPath.Split(@"\".ToCharArray()).ToList<string>();

            SortedList<string, string> notebookIDs = new SortedList<string, string>();

            _oneNote.GetHierarchy(null, HierarchyScope.hsNotebooks, out data);

            XMLHelper.GetNameIDPairs(data, "one:Notebook", ref notebookIDs);

            // walk down the path to find the notebook ID based on name
            for (int i = 0; i < parts.Count; i++)
            {
                if (notebookIDs.Keys.Contains(parts[i]))
                {
                    sectionID = FindSectionByPath(notebookIDs[parts[i]], parts, i + 1);
                    break;
                }
            }

            return sectionID;
        }

        /// <summary>
        /// Gets a path to a section given the section's ID.
        /// </summary>
        /// <param name="id">Internal ID</param>
        /// <returns></returns>
        public string GetSectionPath(string id)
        {
            string data;
            // _oneNote.GetHierarchy(id, HierarchyScope.hsSelf , out data);
            // _oneNote.GetHierarchyParent()
            return string.Empty;
        }

        private string FindSectionByPath(string id, List<string> parts, int partsIndex)
        {

            string data = null;

            SortedList<string, string> children = new SortedList<string, string>();

            _oneNote.GetHierarchy(id, HierarchyScope.hsChildren, out data);

            // one:Section 
            // one:SectionGroup

            if (partsIndex == parts.Count - 1)
                XMLHelper.GetNameIDPairs(data, "one:Section", ref children);
            else
                XMLHelper.GetNameIDPairs(data, "one:SectionGroup", ref children);

            // walk down the path to find the notebook ID based on name
            if (children.Keys.Contains(parts[partsIndex]))
            {
                if (partsIndex == parts.Count - 1)
                    id = children[parts[partsIndex]];
                else
                    id = FindSectionByPath(children[parts[partsIndex]], parts, partsIndex + 1);
            }

            return id;

        }

        #endregion

        #region - misc - 

        /// <summary>
        /// Returns a LinkInfo object based on a hyperlink object.
        /// </summary>
        /// <param name="link"></param>
        /// <returns></returns>
        public LinkInfo GetLinkInfo(Hyperlink link)
        {
            return GetLinkInfo(link.Reference);
        }

        /// <summary>
        /// Returns a LinkInfo object based on a full hyperlink string.
        /// </summary>
        /// <param name="link">Full hyperlink path</param>
        /// <returns></returns>
        public LinkInfo GetLinkInfo(string link)
        {
            return new LinkInfo(link);
        }

        /// <summary>
        /// Returns a LinkInfo object to an OneNote object given its internal ID.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public LinkInfo GetLinkInfoFromItem(string id)
        {
            string linkData = GetLinkToItem(id);
            return new LinkInfo(linkData);
        }

        /// <summary>
        /// Returns a fully encoded hyperlink to a OneNote object given its internal ID.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetLinkToItem(string id)
        {

            string link = string.Empty;

            _oneNote.GetHyperlinkToObject(id, null, out link);

            return link;

        }

        /// <summary>
        /// Returns true if the OneNote client is visible.
        /// </summary>
        public bool AppVisible()
        {
            return _oneNote.Windows.Count != 0;
        }

        /// <summary>
        /// Returns the object type based on the objects internal ID
        /// </summary>
        /// <param name="id">Internal ID of the object</param>
        /// <returns></returns>
        public OneNoteObjType GetTypeByID(string id)
        {

            OneNoteObjType resp = OneNoteObjType.Unknown;

            if (string.IsNullOrEmpty(id))
                return resp;

            string data;

            _oneNote.GetHierarchy(id, HierarchyScope.hsSelf, out data);

            string tag = data.GetInsideValue("<one:", " ", false);
            switch (tag)
            {
                case "Page":
                    resp = OneNoteObjType.Page;
                    break;
                case "Section":
                    resp = OneNoteObjType.Section;
                    break;
                case "SectionGroup":
                    resp = OneNoteObjType.SectionGroup;
                    break;
                case "Notebook":
                    resp = OneNoteObjType.Notebook;
                    break;
                default:
                    break;
            }

            return resp;
        }

        /// <summary>
        /// Gets the parent ID given an object's internal ID.
        /// </summary>
        /// <param name="id">Internal object ID</param>
        /// <returns></returns>
        public string GetParentID(string id)
        {
            string parentID;

            _oneNote.GetHierarchyParent(id, out parentID);

            return parentID;
        }

        /// <summary>
        /// Validates the hyperlink represented by the given LinkInfo object.
        /// </summary>
        /// <param name="linkInfo"></param>
        /// <returns></returns>
        public bool ValidateLink(LinkInfo linkInfo)
        {
            bool resp = false;
            switch (linkInfo.LinkType)
            {
                case LinkInfo.LinkTypeEnum.Unknown:
                    break;
                case LinkInfo.LinkTypeEnum.File:
                    resp = ValidationHelper.ValidateFile(linkInfo);
                    break;
                case LinkInfo.LinkTypeEnum.Directory:
                    resp = ValidationHelper.ValidateDir(linkInfo);
                    break;
                case LinkInfo.LinkTypeEnum.Web:
                    resp = ValidationHelper.ValidateWeb(linkInfo);
                    break;
                case LinkInfo.LinkTypeEnum.MailTo:
                    resp = ValidationHelper.ValidateMailTo(linkInfo);
                    break;
                case LinkInfo.LinkTypeEnum.Page:
                    resp = ValidationHelper.ValidatePage(linkInfo, this);
                    break;
                case LinkInfo.LinkTypeEnum.Section:
                    resp = ValidationHelper.ValidateSection(linkInfo, this);
                    break;
                case LinkInfo.LinkTypeEnum.Group:
                    resp = ValidationHelper.ValidateGroup(linkInfo, this);
                    break;
                case LinkInfo.LinkTypeEnum.Notebook:
                    resp = ValidationHelper.ValidateNotebook(linkInfo, this);
                    break;
                default:
                    break;
            }

            return resp;
        }

        /// <summary>
        /// Validates the hyperlink represented by the given Hyperlink object.
        /// </summary>
        /// <param name="linkInfo"></param>
        /// <returns></returns>
        public bool ValidateLink(Hyperlink hyperlink)
        {
            return ValidateLink(new LinkInfo(hyperlink.Reference));
        }

        /// <summary>
        /// Validates the hyperlink represented by the given hyperlink string.
        /// </summary>
        /// <param name="linkInfo"></param>
        /// <returns></returns>
        public bool ValidateLink(string link)
        {
            return ValidateLink(new LinkInfo(link));
        }

        #endregion

    }

}

