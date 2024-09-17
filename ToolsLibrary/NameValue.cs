namespace OneNoteTools
{
    public class NameValue
    {

        public string Name;
        public string Value;

        public NameValue()
        {
            // nothing
        }

        public NameValue(string name)
        {

            Name = name;

        }

        public NameValue(string nameValuePair, string delimiter, bool isPair)
        {
            string[] parts = nameValuePair.Split(delimiter);
            Name = parts[0].Trim();
            if (parts.Length > 1)
            {
                Value = parts[1].Trim();
            }
        }

        public NameValue(string name, string id)
        {

            Name = name;
            Value = id;

        }
    }
}
