namespace LibVisioSharp.Model
{
    public class Shape
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string NameU { get; set; }    
        public string Type { get; set; }
        public string Text { get; set; }
        public string MasterId { get; set; }
        public string MasterShapeId { get; set; }
        public List<Cell> Cells { get; set; }
        public string LineStyle { get; set; }
        public string FillStyle { get; set; }
        public string TextStyle { get; set; }
        public string ThemeTextColor { get; set; }
        public string ComputedFill { get; set; }
        public float? MasterWidth { get; set; }
        public float? MasterHeight { get; set; }

        public List<Section> Geometries { get; set; }
        public List<Section> Controls { get; set; }
        public List<Section> Connections { get; set; }
        public Section User { get; set; }
        public Section CharacterFormats { get; set; }
        public Section ParagraphFormats { get; set; }
        public List<GradientStop> GradientStops { get; set; }
        public List<TextInfo> TextParts { get; set; }
        public ForeignData ForeignData { get; set; }
        public List<HyperLink> HyperLinks { get; set; }
        public Shape MasterShape { get; set; }
        public List<Shape> SubShapes { get; set; }
        public bool HasOwnGeometry { get; set; }

        public bool HasGeometry => this.Geometries != null && this.Geometries.Count > 0;
        public bool HasConnection => this.Connections != null && this.Connections.Count > 0;
        public bool HasUser => this.User != null;
        public bool HasControl => this.Controls != null && this.Controls.Count > 0;
        public bool HasTextElement => this.Text !=null;
        public bool HasCharacterFormat => this.CharacterFormats != null;
        public bool HasParagraphFormat => this.ParagraphFormats != null;
        public bool HasSubShape => this.SubShapes != null && this.SubShapes.Count > 0;
    }
}
