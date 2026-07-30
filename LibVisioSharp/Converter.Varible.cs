using System;
using System.Collections.Generic;
using System.Text;

namespace LibVisioSharp
{
    public partial class Converter
    {
        // Visio XML namespaces
        private static readonly Dictionary<string, string> _NS = new Dictionary<string, string>()
        {
         { "v", "http://schemas.microsoft.com/office/visio/2012/main" },
         { "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships" }
        };
        private static readonly string _VNS = _NS["v"];
        private static readonly string _VTAG = "{" + _VNS + "}";

        // Supported file extensions
        public static readonly string[] VISIO_EXTENSIONS = [".vsd", ".vsdx", ".vsdm"];
        public static readonly string[] TEMPLATE_EXTENSIONS = [".vst", ".vstx", ".vstm"];
        public static readonly string[] STENCIL_EXTENSIONS = [".vss", ".vssx", ".vssm"];
        public static readonly string[] ALL_EXTENSIONS = Enumerable.Concat(Enumerable.Concat(VISIO_EXTENSIONS, TEMPLATE_EXTENSIONS), STENCIL_EXTENSIONS).ToArray();

        // XML-based (ZIP) formats that use the built-in parser
        public static readonly string[] _XML_EXTENSIONS = [".vsdx", ".vsdm", ".vssx", ".vssm", ".vstx", ".vstm"];



        // Visio line patterns
        private static readonly Dictionary<int, string> _LINE_PATTERNS = new Dictionary<int, string>()
        {
            {0, "none"},           // No line
            {1, ""},               // Solid
            {2, "4,3"},            // Dash
            {3, "1,3"},            // Dot
            {4, "4,3,1,3"},        // Dash-dot
            {5, "4,3,1,3,1,3"},    // Dash-dot-dot
            {6, "8,3"},            // Long dash
            {7, "1,1"},            // Dense dot
            {8, "8,3,1,3"},        // Long dash-dot
            {9, "8,3,1,3,1,3"},   // Long dash-dot-dot
            {10, "12,6"},          // Extra-long dash
            {16, "6,3,6,3"}       // Dash-dash
        };

        // Inches to SVG pixels conversion
        private const float _INCH_TO_PX = 72.0f;

        // Arrow size lookup (BeginArrowSize/EndArrowSize 0-6 -> scale factor)
        private static readonly Dictionary<int, float> _ARROW_SIZES = new Dictionary<int, float>() { { 0, 0.6f }, { 1, 0.8f }, { 2, 1.0f }, { 3, 1.2f }, { 4, 1.6f }, { 5, 2.0f }, { 6, 2.5f } };

        // MIME types for embedded images
        private static readonly Dictionary<string, string> _IMAGE_MIMETYPES = new Dictionary<string, string>()
        {
            { ".png", "image/png"},
            {".jpg", "image/jpeg"},
            {".jpeg", "image/jpeg"},
            {".gif", "image/gif"},
            {".bmp", "image/bmp"},
            {".emf", "image/x-emf"},
            {".wmf", "image/x-wmf"},
            {".tiff", "image/tiff"},
            {".tif", "image/tiff"},
            {".svg", "image/svg+xml"}
        };

        // Relationship namespace
        private const string _RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

        // Font family mapping: Visio font names -> SVG-compatible font stacks
        private static readonly Dictionary<string, string> _FONT_MAP = new Dictionary<string, string>()
        {
            {"angsana new", "Noto Sans Thai}, Noto Serif Thai}, sans-serif"},
            {"browallia new", "Noto Sans Thai}, sans-serif"},
            {"cordia new", "Noto Sans Thai}, sans-serif"},
            {"freesia upc", "Noto Sans Thai}, sans-serif"},
            {"tahoma", "Tahoma}, Noto Sans}, sans-serif"},
            {"arial", "Arial}, Noto Sans}, sans-serif"},
            {"calibri", "Calibri}, Noto Sans}, sans-serif"},
            {"segoe ui", "Segoe UI}, Noto Sans}, sans-serif"},
            {"times new roman", "Times New Roman}, Noto Serif}, serif"},
            {"ms gothic", "Noto Sans JP}, sans-serif"},
            {"ms mincho", "Noto Serif JP}, serif"},
            {"simsun", "Noto Sans SC}, sans-serif"},
            {"simhei", "Noto Sans SC}, sans-serif"},
            {"microsoft yahei", "Noto Sans SC}, sans-serif"},
            {"malgun gothic", "Noto Sans KR}, sans-serif"},
            {"gulim", "Noto Sans KR}, sans-serif"}
        };
    }
}
