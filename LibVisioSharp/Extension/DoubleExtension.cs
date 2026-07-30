namespace LibVisioSharp.Extension
{
    public static class DoubleExtension
    {
        public static string ToFixed(this double value, int number)
        {
            return value.ToString("0." + string.Join("", Enumerable.Repeat("0", number)));
        }       
    }
}
