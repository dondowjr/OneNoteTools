using System.Net;
using System.Net.NetworkInformation;

namespace OneNoteTools
{
    /// <summary>
    /// This class contains helper routines for validating OneNote hyperlinks.
    /// </summary>
    public static class ValidationHelper
    {
        private static SortedList<string, bool> _resultCache = new SortedList<string, bool>();

        public static void ClearCache()
        {
            _resultCache.Clear();
        }

        public static bool CheckCache(string path, out bool evalResult)
        {

            bool resp = false;
            evalResult = false;

            if (_resultCache.ContainsKey(path))
            {
                resp = true;
                evalResult = _resultCache[path];
            }

            return resp;

        }

        public static void AddToCache(string path, bool evalResult)
        {

            //System.Threading.Thread.Yield();
            lock (_resultCache)
            {
                if (!_resultCache.ContainsKey(path))
                    _resultCache.Add(path, evalResult);
            }
        }

        public static bool ValidateFile(LinkInfo info)
        {
            bool resp = false;

            if (CheckCache("file:" + info.FullPathPlainText, out resp))
                return resp;

            try
            {
                return File.Exists(info.FullPathPlainText);
            }
            catch (Exception) { }

            AddToCache("file:" + info.FullPathPlainText, resp);

            return resp;
        }

        public static bool ValidateDir(LinkInfo info)
        {
            bool resp = false;

            if (CheckCache("dir:" + info.FullPathPlainText, out resp))
                return resp;

            try
            {
                resp = Directory.Exists(info.FullPathPlainText);
            }
            catch (Exception) { }

            AddToCache("dir:" + info.FullPathPlainText, resp);

            return resp;
        }

        public static bool ValidateWeb(LinkInfo info)
        {

            bool resp = false;
            string address = info.FullPath;
            string original = address;

            // validate the full url
            if (CheckCache("webfull:" + address, out resp))
                return resp;

            if (ValidateURL(address))
            {
                AddToCache("webfull:" + address, true);
                return true;
            }

            // validate url w/o arguments
            int s = -1;
            do
            {
                s = address.IndexOf(".", s + 1);

                if (s == -1 || s + 4 >= address.Length)
                {
                    s = -1;
                    break;
                }

                if (address.Substring(s + 4, 1) == "/")
                {
                    address = address.Substring(0, s + 4);
                    break;
                }

            } while (s != -1);

            if (s > -1)
            {

                if (CheckCache("webnoarg:" + address, out resp))
                    return resp;

                if (ValidateURL(address))
                {
                    AddToCache("webfull:" + original, resp);
                    AddToCache("webnoarg:" + address, true);
                    return true;
                }
            }

            // https to http
            if (address.Substring(0, 8) == "https://")
                address = "http://" + address.Substring(8);

            if (CheckCache("webhttp:" + address, out resp))
                return resp;

            if (ValidateURL(address))
            {
                AddToCache("webfull:" + original, resp);
                AddToCache("webhttp:" + address, true);
                return true;
            }

            // validate url w/o the http or https prefix
            if (address.Substring(0, 7) == "http://")
                address = address.Substring(7);

            if (CheckCache("webnohttp:" + address, out resp))
                return resp;

            if (ValidateURL(address))
            {
                AddToCache("webfull:" + original, resp);
                AddToCache("webnohttp:" + address, true);
                return true;
            }

            // try with www
            address = "www." + address;

            if (CheckCache("webwww:" + address, out resp))
                return resp;

            resp = ValidateURL(address);

            AddToCache("webfull:" + original, resp);

            return resp;

        }

        public static bool PingWebSite(string address)
        {

            bool resp = false;

            if (CheckCache("ping:" + address, out resp))
                return resp;

            try
            {
                using (Ping ping = new Ping())
                {
                    PingReply pResp = ping.Send(address);
                    resp = (pResp.Status == IPStatus.Success);
                }
            }
            catch (Exception) { }

            AddToCache("ping:" + address, resp);

            return resp;
        }

        public static string GetDomainFromAddress(string address)
        {
            // strip arguments
            int s = -1;
            do
            {
                s = address.IndexOf(".", s + 1);

                if (s == -1 || s + 4 >= address.Length)
                {
                    s = -1;
                    break;
                }

                if (address.Substring(s + 4, 1) == "/")
                {
                    address = address.Substring(0, s + 4);
                    break;
                }

            } while (s != -1);

            // strip http
            if (address.Length >= 8)
            {
                if (address.Substring(0, 8) == "https://")
                    address = address.Substring(8);
            }

            if (address.Length >= 7)
            {
                if (address.Substring(0, 7) == "http://")
                    address = address.Substring(7);
            }

            return address;

        }

        private static bool ValidateURL(string address)
        {
            bool resp = false;

            try
            {

                Uri url = new Uri(address);
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Timeout = 2000;
                req.UseDefaultCredentials = true;
                req.Method = "HEAD";
                req.AllowAutoRedirect = true;

                using (HttpWebResponse rep = (HttpWebResponse)req.GetResponse())
                {
                    rep.Close();
                    resp = true;
                }
            }
            catch (Exception ex)
            {
                resp = false;
            }

            return resp;
        }

        public static bool ValidateMailTo(LinkInfo info)
        {

            int i = -1;
            string address = info.FullPathPlainText;

            // peal off mailto:
            if (address.Substring(0, 7) != "mailto:")
                return false;
            address = address.Substring(7);

            // parse off arguments (if they exists)
            i = address.IndexOf("?");
            if (i != -1)
                address = address.Substring(0, i);

            // parse off the name
            i = address.IndexOf("@");
            if (i == -1)
                return false;

            address = address.Substring(i + 1);

            return PingWebSite(address);

        }

        public static bool ValidatePage(LinkInfo info, Connection conn)
        {
            bool resp = false;

            if (CheckCache("page:" + info.FullPathPlainText, out resp))
                return resp;

            string sectionID = string.Empty;

            if (!string.IsNullOrEmpty(info.FullPathPlainText) && !string.IsNullOrEmpty(info.PageName))
            {
                sectionID = conn.GetSectionID(info.FullPathPlainText);
                if (!string.IsNullOrEmpty(sectionID))
                {
                    foreach (PageInfo item in conn.GetPagesInfo(sectionID))
                    {
                        if (item.Name == info.PageName)
                        {
                            resp = true;
                            break;
                        }
                    }
                }
            }

            AddToCache("page:" + info.FullPathPlainText, resp);

            return resp;

        }

        public static bool ValidateSection(LinkInfo info, Connection conn)
        {
            bool resp = false;

            if (CheckCache("section:" + info.FullPathPlainText, out resp))
                return resp;

            if (info.ExternalLink)
            {
                if (!string.IsNullOrEmpty(info.FullPathPlainText))
                {
                    string id = conn.GetSectionID(info.FullPathPlainText);
                    resp = !string.IsNullOrEmpty(id);
                }
            }
            else
            {
                resp = ValidateFile(info);
            }

            AddToCache("section:" + info.FullPathPlainText, resp);

            return resp;
        }

        public static bool ValidateGroup(LinkInfo info, Connection conn)
        {

            bool resp = false;

            if (CheckCache("group:" + info.FullPathPlainText, out resp))
                return resp;

            if (info.ExternalLink)
            {
                throw new NotImplementedException();
            }
            else
            {
                resp = ValidateDir(info);
            }

            AddToCache("group:" + info.FullPathPlainText, resp);

            return resp;
        }

        public static bool ValidateNotebook(LinkInfo info, Connection conn)
        {
            bool resp = false;

            if (CheckCache("notebook:" + info.FullPathPlainText, out resp))
                return resp;

            //throw new NotImplementedException();

            AddToCache("notebook:" + info.FullPathPlainText, resp);

            return resp;

        }

    }
}
