namespace LibVisioNgSharp.Model
{
    public class Document
    {
        public List<StyleSheet> StyleSheets { get; set; }
        public Dictionary<string, string> Colors { get; set; }

        public bool HasStyleSheet => this.StyleSheets != null && this.StyleSheets.Count > 0;
        public bool HasColor => this.Colors != null && this.Colors.Count > 0;
    }
}
