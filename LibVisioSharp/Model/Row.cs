namespace LibVisioSharp.Model
{
    public class Row
    {
        public string Index { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public bool IsDelete { get; set; }
        public List<Cell> Cells { get; set; }
    }
}
