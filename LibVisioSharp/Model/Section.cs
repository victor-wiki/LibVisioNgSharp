namespace LibVisioSharp.Model
{
    public class Section
    {
        public string Index { get; set; }
        public string Name { get; set; }
        public List<Cell> Cells { get; set; }
        public List<Row> Rows { get; set; }
        public bool NoFill { get; set; }
        public bool NoLine { get; set; }
        public bool NoShow { get; set; }

        public bool HasRow => this.Rows != null && this.Rows.Count > 0;
        public bool HasCell => this.Cells != null && this.Cells.Count > 0;
    }
}
