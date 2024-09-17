namespace OneNoteTools
{
    /// <summary>
    /// This class contains extension methods.
    /// </summary>
    public static class Extensions
    {
        public static List<string> GetInsideValues(this string value, string startTag, string endTag, bool incTags)
        {

            int s = 0;
            int e = 0;
            List<string> items = new List<string>();
            string temp = string.Empty;

            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(startTag) || string.IsNullOrEmpty(endTag))
                return items;

            while (true)
            {
                s = value.IndexOf(startTag, s);
                if (s > -1)
                {
                    s += startTag.Length;
                    e = value.IndexOf(endTag, s);
                    if (e > -1)
                    {
                        temp = value.Substring(s, e - s);

                        if (temp.Length > 0)
                        {
                            if (incTags)
                                temp = startTag + temp + endTag;

                            items.Add(temp);

                        }
                        s = e + endTag.Length;
                    }
                    else
                        break;
                }
                else
                    break;
            }
            return items;

        }

        public static string GetInsideValue(this string value, string startTag, string endTag)
        {

            return GetInsideValue(value, startTag, endTag, false, false);

        }

        public static string GetInsideValue(this string value, string startTag, string endTag, bool incTags)
        {

            return GetInsideValue(value, startTag, endTag, incTags, false);

        }

        public static string GetInsideValue(this string value, string startTag, string endTag, bool includeTags, bool fromEnd)
        {

            int s = 0;
            int e = 0;
            string temp = string.Empty;

            if (string.IsNullOrEmpty(value))
                return temp;

            if (string.IsNullOrEmpty(startTag) && string.IsNullOrEmpty(endTag))
                return temp;

            if (string.IsNullOrEmpty(startTag))
                s = 0;
            else
            {
                if (fromEnd)
                    s = value.LastIndexOf(startTag);
                else
                    s = value.IndexOf(startTag);
            }

            if (s > -1)
            {
                s += startTag.Length;

                if (string.IsNullOrEmpty(endTag))
                    e = value.Length;
                else
                {
                    if (fromEnd)
                    {
                        if (string.IsNullOrEmpty(startTag))
                            e = value.LastIndexOf(endTag);
                        else
                            e = value.IndexOf(endTag, s);
                    }
                    else
                        e = value.IndexOf(endTag, s);
                }

                if (e > -1)
                {
                    temp = value.Substring(s, e - s);

                    if (includeTags)
                        temp = startTag + temp + endTag;

                }
            }

            return temp;

        }

        public static string Decode(this string value)
        {

            return Uri.UnescapeDataString(value);

        }

        public static string Right(this string value, int chars)
        {

            if (!string.IsNullOrEmpty(value))
            {
                if (value.Length >= chars)
                    value = value.Substring(value.Length - chars);
            }

            return value;

        }

        public static string Left(this string value, int chars)
        {

            if (!string.IsNullOrEmpty(value))
            {
                if (value.Length >= chars)
                    value = value.Substring(0, chars);
            }

            return value;

        }

        public static string CropRight(this string value, int chars)
        {

            if (!string.IsNullOrEmpty(value))
            {
                if (value.Length > chars)
                    value = value.Substring(0, value.Length - chars);
            }

            return value;


        }

        // todo: delete this routine?
        public static string CropLeft(this string value, int chars)
        {

            if (!string.IsNullOrEmpty(value))
            {
                if (value.Length > chars)
                    value = value.Substring(chars);

            }

            return value;

        }


        public static string[] Split(this string value, string separator)
        {

            return value.Split(separator.ToCharArray());

        }

    }
}
