using Newtonsoft.Json;

namespace LibVisioNgSharp.Helper
{
    public class ObjectHelper
    {
        public static T CloneObject<T>(object obj)
        {
            return (T)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(obj), typeof(T));
        }


        public static object GetValue(object obj, string propertyName)
        {
            if (obj == null || propertyName == null)
            {
                return null;
            }

            var property = obj.GetType().GetProperties().FirstOrDefault(item => item.Name == propertyName);

            if (property != null)
            {
                return property.GetValue(obj);
            }

            return null;
        }
    }
}
