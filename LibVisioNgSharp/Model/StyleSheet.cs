namespace LibVisioNgSharp.Model
{
    public class StyleSheet
    {
        public string Id { get; set; }
        public List<Cell> Cells { get; set; }
        public string LineStyle { get; set; }
        public string FillStyle { get; set; }
        public string TextStyle { get; set; }
    }
}
