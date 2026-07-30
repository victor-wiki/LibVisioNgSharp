using System.Xml.Linq;

namespace LibVisioSharp.Extension
{
    public static class XElementExtension
    {
        public static string GetAttributeValue(this XElement element, string name, string defaultValue = null)
        {
            if (element == null)
            {
                return defaultValue;
            }

            string localName = null;

            if(name.Contains(":"))
            {
                localName = name.Split(':').Last();
            }
            else
            {
                localName = name;
            }

            foreach (var attr in element.Attributes())
            {
                if(attr.Name.LocalName == localName)
                {
                    return attr.Value?? defaultValue;
                }
            }             

            return defaultValue;
        }

        public static XElement Child(this XElement element, string name)
        {
            if (element == null)
            {
                return null;
            }

            return element.Elements().FirstOrDefault(item => item.Name.LocalName == name);
        }

        public static List<XElement> Children(this XElement element, string name)
        {
            if (element == null)
            {
                return null;
            }

            return element.Elements().Where(item => item.Name.LocalName == name).ToList();
        }

        public static bool HasChild(this XElement element, string name)
        {
            if (element == null)
            {
                return false;
            }

            return element.Elements().FirstOrDefault(item => item.Name.LocalName == name) != null;
        }

        public static long? GetNumberValue(this XElement element, string name)
        {
            if (element == null)
            {
                return null;
            }

            var value = element.Attribute(name)?.Value;

            if (long.TryParse(value, out var val))
            {
                return val;
            }

            return null;
        }
    }
}
