using ICSharpCode.SharpZipLib.Core;
using ImageMagick;
using LibVisioSharp.Extension;
using LibVisioSharp.Helper;
using LibVisioSharp.Model;
using NaturalSort.Extension;
using NCalc;
using PowerPointConverter.Extension;
using SharpCompress.Archives;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace LibVisioSharp
{
    public partial class Converter
    {
        /// <summary>
        /// Get SVG stroke-dasharray for a Visio line pattern.
        /// </summary>
        /// <param name="pattern"></param>
        /// <param name="weight"></param>
        /// <returns></returns>
        private static string _get_dash_array(int pattern, float weight)
        {
            if (pattern == 0)
            {
                return "none";
            }

            var p = _LINE_PATTERNS.ContainsKey(pattern) ? _LINE_PATTERNS[pattern] : null;

            if (string.IsNullOrEmpty(p) || p == "none")
            {
                // For unknown patterns 2-23, generate a reasonable dash pattern
                if (pattern >= 2 && pattern <= 23)
                {
                    // Generate based on pattern number
                    if (pattern % 3 == 0)
                    {
                        p = "1,2";  // dot-like
                    }
                    else if (pattern % 3 == 1)
                    {
                        p = "6,3";  // dash-like
                    }
                    else
                    {
                        p = "6,3,1,3";  // dash-dot
                    }
                }
                else
                {
                    return null;
                }
            }

            // Scale dash pattern by stroke weight
            var scale = Math.Max(weight, 0.5);

            var items = p.Split(',');
            var parts = items.Select(item => Math.Round(float.Parse(item) * scale, 1));

            return string.Join(",", parts);
        }

        /// <summary>
        /// Escape text for XML/SVG output.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static string _escape_xml(string text)
        {
            return text.Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&apos;");
        }

        #region Embedded image support

        /// <summary>
        /// Parse theme colors from visio/theme/theme1.xml.
        /// </summary>
        /// <param name="zf"></param>
        /// <returns>Returns a dict mapping theme color names to #RRGGBB values.
        /// Keys: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
        /// Also maps numeric indices used by Visio THEMEVAL:
        /// 0->dk1, 1->lt1, 2->dk2, 3->lt2, 4->accent1, ..., 9->accent6,
        /// 10->hlink, 11->folHlink
        /// </returns>
        private static Dictionary<string, string> _parse_theme(IArchive zf)
        {
            var _DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

            var theme_colors = new Dictionary<string, string>();

            string[] paths = ["visio/theme/theme1.xml", "visio/theme/theme2.xml"];

            foreach (string path in paths)
            {
                string theme_xml = GetFileContent(zf, path);

                if (theme_xml == null)
                {
                    continue;
                }

                XElement root = XDocument.Parse(theme_xml).Root;

                var elements = root.Child("themeElements").Child("clrScheme").Elements();

                foreach (var el in elements)
                {
                    string name = el.Name.LocalName;

                    var srgbClrElement = el.Child("srgbClr");

                    string value = null;

                    if (srgbClrElement != null)
                    {
                        value = srgbClrElement?.GetAttributeValue("val");
                    }
                    else
                    {
                        var sysClrElement = el.Child("sysClr");

                        if (sysClrElement != null)
                        {
                            value = sysClrElement.GetAttributeValue("val");
                        }
                    }

                    if (value != null && value.Length == 6)
                    {
                        theme_colors.Add(name, "#" + value);
                    }
                }

                if (theme_colors != null && theme_colors.Count > 0)
                {
                    break;
                }
            }

            // Build numeric index mapping (Visio theme color indices)
            var _idx_map = new Dictionary<int, string>()
            {
                { 0, "dk1" }, {1, "lt1" }, {2, "dk2" }, {3, "lt2" },
                { 4, "accent1" }, {5, "accent2" }, {6, "accent3" }, {7, "accent4" },
                { 8, "accent5" }, {9, "accent6" }, {10, "hlink" }, {11, "folHlink" },
            };

            foreach (var item in _idx_map)
            {
                if (theme_colors.ContainsKey(item.Value.ToString()))
                {
                    theme_colors.Add(item.Key.ToString(), theme_colors[item.Value]);
                }
            }

            return theme_colors;
        }

        private static string GetFileContent(IArchive zf, string path)
        {
            var entry = zf.Entries.FirstOrDefault(item => item.Key == path);

            if (entry == null)
            {
                return null;
            }

            using Stream zipStream = entry.OpenEntryStream();

            MemoryStream ms = ConvertToMemoryStream(zipStream);

            using (StreamReader reader = new StreamReader(ms))
            {
                return reader.ReadToEnd();
            }
        }

        private static byte[] GetFileBytes(IArchive zf, string path)
        {
            var entry = zf.Entries.FirstOrDefault(item => item.Key == path);

            using Stream zipStream = entry.OpenEntryStream();

            MemoryStream ms = ConvertToMemoryStream(zipStream);

            return ms.ToArray();
        }

        private static MemoryStream ConvertToMemoryStream(Stream stream)
        {
            MemoryStream memoryStream = new MemoryStream();

            byte[] buffer = new byte[4096];

            StreamUtils.Copy(stream, memoryStream, buffer);

            memoryStream.Position = 0;

            return memoryStream;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="page_xml_root"></param>
        /// <returns>{layer_index: {"name": str, "visible": bool}}</returns>
        public static Dictionary<string, Layer> _parse_layers(XElement page_xml_root)
        {
            var layers = new Dictionary<string, Layer>();

            var page_sheet = page_xml_root.Child("Page")?.Child("PageSheet");

            if (page_sheet == null)
            {
                return layers;
            }

            var sections = page_sheet.Children("Section");

            foreach (var section in sections)
            {
                if (section.GetAttributeValue("N") == "Layer")
                {
                    continue;
                }

                var rows = section.Children("Row");

                foreach (var row in rows)
                {
                    var ix = row.GetAttributeValue("IX");

                    Dictionary<string, string> dictCells = new Dictionary<string, string>();

                    foreach (var cell in row.Children("Cell"))
                    {
                        string name = cell.GetAttributeValue("N");
                        string value = cell.GetAttributeValue("V");

                        dictCells.Add(name, value);
                    }

                    layers.Add(ix, new Layer() { Name = dictCells["Name"], Visible = dictCells["Visible"] == "1" });
                }
            }

            return layers;
        }

        /// <summary>
        /// Parse visio/document.xml
        /// </summary>
        /// <param name="zf"></param>
        /// <returns></returns>
        private static Document _parse_document(IArchive zf)
        {
            Document doc = new Document();

            string doc_xml = GetFileContent(zf, "visio/document.xml");

            var root = XDocument.Parse(doc_xml).Root;

            var styleSheets = root.Child("StyleSheets")?.Children("StyleSheet");

            if (styleSheets != null)
            {
                doc.StyleSheets = new List<StyleSheet>();

                foreach (var ss in styleSheets)
                {
                    string sid = ss.GetAttributeValue("ID");

                    if (string.IsNullOrEmpty(sid))
                    {
                        continue;
                    }

                    var cells = new List<Cell>();

                    var cellElements = ss.Children("Cell");

                    foreach (var cell in cellElements)
                    {
                        var n = cell.GetAttributeValue("N");
                        var v = cell.GetAttributeValue("V");
                        var f_attr = cell.GetAttributeValue("F");

                        cells.Add(new Cell { Name = n, Value = v, Formula = f_attr });
                    }

                    StyleSheet styleSheet = new StyleSheet()
                    {
                        Id = sid,
                        Cells = cells,
                        LineStyle = ss.GetAttributeValue("LineStyle"),
                        FillStyle = ss.GetAttributeValue("FillStyle"),
                        TextStyle = ss.GetAttributeValue("TextStyle")
                    };

                    doc.StyleSheets.Add(styleSheet);
                }
            }

            var colors = root.Child("Colors")?.Children("ColorEntry");

            if (colors != null)
            {
                doc.Colors = new Dictionary<string, string>();

                foreach (var color in colors)
                {
                    doc.Colors.Add(color.GetAttributeValue("IX"), color.GetAttributeValue("RGB"));
                }
            }

            return doc;
        }

        /// <summary>
        /// Walk the StyleSheet inheritance chain to resolve a cell value.
        /// </summary>
        /// <param name="styles"></param>
        /// <param name="style_id"></param>
        /// <param name="cell_name"></param>
        /// <param name="category"></param>
        /// <param name="_depth"></param>
        /// <returns></returns>
        private static string _resolve_style_cell(Dictionary<string, dynamic> styles, string style_id, string cell_name, string category = "line", int _depth = 0)
        {
            if (_depth > 10 || style_id == null || !styles.ContainsKey(style_id))
            {
                return null;
            }

            var ss = styles[style_id];
            var cell = ss["cells"][cell_name];
            var val = cell.V;
            var formula = cell.F;

            if (!string.IsNullOrEmpty(val) && formula != "Inh" && val != "Themed")
            {
                return val;
            }

            var parent_keys = new Dictionary<string, string>(){
                { "line", "line_style" }, {"fill", "fill_style" },
                { "text", "text_style"}};

            var parent_key = parent_keys.ContainsKey(category) ? parent_keys[category] : "line_style";

            var parent_id = ObjectHelper.GetValue(ss, parent_key);

            if (!string.IsNullOrEmpty(parent_id) && parent_id != style_id)
            {
                return _resolve_style_cell(styles, parent_id, cell_name, category, _depth + 1);
            }

            return val;
        }

        /// <summary>
        /// Extract all files from visio/media/ in the ZIP.
        /// 
        /// Returns {filename: bytes} e.g. {"image1.png": b"..."}
        /// </summary>
        /// <param name="zf"></param>
        /// <returns></returns>
        private static Dictionary<string, byte[]> _extract_media(IArchive zf)
        {
            var media = new Dictionary<string, byte[]>();

            var keys = zf.Entries.Select(item => item.Key);

            foreach (var key in keys)
            {
                if (key.StartsWith("visio/media/"))
                {
                    string fname = Path.GetFileName(key);

                    media.Add(fname, GetFileBytes(zf, key));
                }
            }

            return media;
        }

        /// <summary>
        /// Parse relationship file for a page to map rId -> target path.
        /// 
        /// For visio/pages/page1.xml, the rels file is
        /// visio/pages/_rels/page1.xml.rels
        /// </summary>
        /// <param name="zf"></param>
        /// <param name="page_file"></param>
        /// <returns></returns>
        private static Dictionary<string, string> _parse_rels(IArchive zf, string page_file)
        {
            var page_dir = Path.GetDirectoryName(page_file).Replace("\\", "/");
            var page_basename = Path.GetFileName(page_file);
            var rels_path = $"{page_dir}/_rels/{page_basename}.rels";

            var rels = new Dictionary<string, string>();

            var rels_xml = GetFileContent(zf, rels_path);

            if (rels_xml == null)
            {
                return null;
            }

            var root = XDocument.Parse(rels_xml).Root;

            foreach (var rel in root.Children("Relationship"))
            {
                var rid = rel.GetAttributeValue("Id");
                var target = rel.GetAttributeValue("Target");

                rels.Add(rid, target);
            }

            return rels;
        }

        /// <summary>
        /// Convert EMF/WMF data to SVG. Returns SVG bytes or None.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="ext"></param>
        /// <returns></returns>
        private static byte[] _convert_emf_to_svg(byte[] data, string ext)
        {
            using (var image = new MagickImage(data))
            {
                image.Format = MagickFormat.Svg;
                image.Quality = 100;

                return image.ToByteArray();
            }
        }

        /// <summary>
        /// onvert image bytes to a base64 data URI.
        /// </summary>
        /// <param name="data"></param>
        /// <param name="filename"></param>
        /// <returns></returns>
        private static string _image_to_data_uri(byte[] data, string filename)
        {
            var ext = Path.GetExtension(filename).ToLower();

            //Convert BMP to PNG for data URI (BMP not widely supported in SVG)
            if (ext == ".bmp" || ext == ".dib")
            {
                using (var image = new MagickImage(data))
                {
                    image.Format = MagickFormat.Png;

                    data = image.ToByteArray();
                }
            }

            if (ext == ".emf" || ext == ".wmf")
            {
                var svg_data = _convert_emf_to_svg(data, ext);

                if (svg_data != null)
                {
                    var b64 = Convert.ToBase64String(svg_data);

                    return $"data:image/svg+xml;base64,{b64}";
                }
                else
                {
                    return null;
                }
            }

            string extName = ext.Trim(',');

            var mime = FileHelper.MimeMappings.ContainsKey(extName) ? FileHelper.MimeMappings[extName] : "image/png";

            var base64 = Convert.ToBase64String(data);

            return $"data:{mime};base64,{base64}";
        }

        /// <summary>
        /// Parse ForeignData element from a shape.
        /// </summary>
        /// <param name="shape_elem"></param>
        /// <returns>Returns {"type": "bitmap"|"metafile", "data": base64_str, "rel_id": rIdN}
        /// or None if no foreign data.
        /// </returns>
        private static ForeignData _parse_foreign_data(XElement shape_elem)
        {
            var fd = shape_elem.Child("ForeignData");

            if (fd == null)
            {
                return null;
            }

            var info = new ForeignData()
            {
                ForeignType = fd.GetAttributeValue("ForeignType"),
                CompressionType = fd.GetAttributeValue("CompressionType")
            };

            // Check for Rel element (can be in Visio namespace or r: namespace)
            var rel_elem = fd.Child("Rel");

            if (rel_elem != null)
            {
                // The r:id attribute may use full namespace
                info.RelId = rel_elem.GetAttributeValue("id");
            }
            else
            {
                // Inline data
                var text = fd.Value;

                if (!string.IsNullOrEmpty(text.Trim()))
                {
                    info.Data = text.Trim();
                }
            }

            return info;
        }

        #endregion

        #region Arrow marker SVG generation

        /// <summary>
        /// Generate SVG <defs> for arrow markers.
        /// </summary>
        /// <param name="used_markers">set of marker IDs like "arrow_end_3", "arrow_start_2"</param>
        /// <returns></returns>
        private static List<string> _arrow_marker_defs(HashSet<string> used_markers)
        {
            if (used_markers == null)
            {
                return new List<string>();
            }

            List<string> lines = ["<defs>"];

            foreach (var marker_id in used_markers.OrderBy(item => item, StringComparison.OrdinalIgnoreCase.WithNaturalSort()))
            {
                // Parse: arrow_{start|end}_{size}_{color}
                var parts = marker_id.Split('_');

                string direction = parts.Length > 1 ? parts[1] : "end";
                var size_idx = parts.Length > 2 ? int.Parse(parts[2]) : 3;
                var color = parts.Length > 3 ? $"#{parts[3]}" : "#333333";

                var scale = _ARROW_SIZES.ContainsKey(size_idx) ? _ARROW_SIZES[size_idx] : 1.0f;
                float marker_w = 10 * scale;
                float marker_h = 7 * scale;

                StringBuilder sb = new StringBuilder();

                if (direction == "start")
                {
                    // Reverse triangle for start                   
                    sb.Append($"<marker id=\"{marker_id}\" markerWidth=\"{marker_w:.1f}\" ");
                    sb.Append($"markerHeight=\"{marker_h.ToFixed(1)}\" refX=\"0\" refY=\"{marker_h / 2:.1f}\" ");
                    sb.Append("orient=\"auto\" markerUnits=\"userSpaceOnUse\">");
                    sb.Append($"<polygon points=\"{marker_w.ToFixed(1)} 0, 0 {(marker_h / 2).ToFixed(1)}, ");
                    sb.Append($"{marker_w.ToFixed(1)} {marker_h.ToFixed(1)}\" fill = \"{color}\" /> ");
                    sb.Append("</marker>");
                }
                else
                {
                    //Forward triangle for end
                    sb.Append($"<marker id=\"{marker_id}\" markerWidth=\"{marker_w.ToFixed(1)}\" ");
                    sb.Append($"markerHeight=\"{marker_h.ToFixed(1)}\" refX=\"{marker_w.ToFixed(1)}\" ");
                    sb.Append($"refY=\"{(marker_h / 2).ToFixed(1)}\" orient=\"auto\" markerUnits=\"userSpaceOnUse\">");
                    sb.Append($"<polygon points=\"0 0, {marker_w.ToFixed(1)} {(marker_h / 2.0f).ToFixed(1)}, ");
                    sb.Append($"0 {marker_h.ToFixed(1)}\" fill=\"{color}\"/>");
                    sb.Append("</marker>");
                }

                lines.Add(sb.ToString());
            }

            lines.Add("</defs>");

            return lines;
        }

        /// <summary>
        /// Generate SVG <defs> for hatching/crosshatch fill patterns.
        /// </summary>
        /// <param name="patterns">{pattern_id: {"fg": color, "bg": color, "type": int}}</param>
        /// <returns></returns>

        /// <summary>
        /// Generate SVG <defs> for gradient fills.
        /// </summary>
        /// <param name="gradients">{grad_id: {"start": color, "end": color, "dir": angle_deg}}</param>
        /// <returns></returns>
        private static List<string> _gradient_defs(Dictionary<string, Gradient> gradients)
        {
            if (gradients == null)
            {
                return new List<string>();
            }

            var lines = new List<string>();

            foreach (var item in gradients.OrderBy(item => item.Key, StringComparison.OrdinalIgnoreCase.WithNaturalSort()))
            {
                var gid = item.Key;
                var g = item.Value;

                StringBuilder sb = new StringBuilder();

                if (g.IsRadial)
                {
                    sb.Append($"<radialGradient id=\"{gid}\" cx=\"50%\" cy=\"50%\" r=\"50%\">");
                    sb.Append($"<stop offset=\"0%\" stop-color=\"{g.StartColor}\"/>");
                    sb.Append($"<stop offset=\"100%\" stop-color=\"{g.StopColor}\"/>");
                    sb.Append("</radialGradient>");
                }
                else
                {
                    var angle = g.Angle;

                    // Convert angle to x1,y1,x2,y2 for linearGradient
                    var rad = angle * Math.PI / 180.0;
                    var x1 = 50 - 50 * Math.Cos(rad);
                    var y1 = 50 + 50 * Math.Sin(rad);
                    var x2 = 50 + 50 * Math.Cos(rad);
                    var y2 = 50 - 50 * Math.Sin(rad);

                    sb.Append($"<linearGradient id=\"{gid}\" ");
                    sb.Append($"x1=\"{x1.ToFixed(1)}%\" y1=\"{y1.ToFixed(1)}%\" x2=\"{x2:.1f}%\" y2=\"{y2.ToFixed(1)}%\">");
                    sb.Append($"<stop offset=\"0%\" stop-color=\"{g.StartColor}\"/>");
                    sb.Append($"<stop offset=\"100%\" stop-color=\"{g.StopColor}\"/>");
                    sb.Append("</linearGradient>");
                }

                lines.Append(sb.ToString());
            }

            return lines;
        }

        /// <summary>
        /// Return SVG filter definition for drop shadows.
        /// </summary>
        /// <returns></returns>
        private static string _shadow_filter_def()
        {
            return "<filter id=\"shadow\" x=\"-10%\" y=\"-10%\" width=\"130%\" height=\"130%\"><feDropShadow dx=\"2\" dy=\"2\" stdDeviation=\"1.5\" flood-color=\"#00000040\"/></filter>";
        }

        #endregion

        #region Master shape parsing

        /// <summary>
        /// Parse full shape data from master files.
        /// </summary>
        /// <param name="zf"></param>
        /// <returns>Returns {master_id: {shape_id: shape_dict, ...}, ...}
        /// Each shape_dict has: cells, geometry, text, char_formats, para_formats, sub_shapes
        /// </returns>
        private static Dictionary<string, Dictionary<string, Shape>> _parse_master_shapes(IArchive zf)
        {
            //First, read masters.xml to map Master ID -> rel ID,
            // then masters.xml.rels to map rel ID -> master file.
            var master_id_to_file = new Dictionary<string, string>();  // Master ID -> master file number

            var master_xml = GetFileContent(zf, "visio/masters/masters.xml");

            var masters = new Dictionary<string, Dictionary<string, Shape>>();

            if (string.IsNullOrEmpty(master_xml))
            {
                return masters;
            }

            var root = XDocument.Parse(master_xml).Root;

            var rid_to_file = new Dictionary<string, string>();

            var rels_xml = GetFileContent(zf, "visio/masters/_rels/masters.xml.rels");
            var rels_root = XDocument.Parse(rels_xml).Root;

            foreach (var rel in rels_root.Children("Relationship"))
            {
                var rid = rel.GetAttributeValue("Id");
                var target = rel.GetAttributeValue("Target");

                // target is like "master2.xml"
                var fname = Path.GetFileNameWithoutExtension(target).Replace("master", "");

                rid_to_file.Add(rid, fname);
            }

            foreach (var master_el in root.Children("Master"))
            {
                var mid = master_el.GetAttributeValue("ID");

                //Find the Rel element — it's in the Visio namespace, not the rels namespace
                var rel_el = master_el.Child("Rel");

                //The r:id attribute uses the relationships namespace
                var rid = rel_el.GetAttributeValue("id");

                if (!string.IsNullOrEmpty(rid) && rid_to_file.ContainsKey(rid))
                {
                    master_id_to_file[mid] = rid_to_file[rid];
                    continue;
                }

                //Fallback: assume master ID matches file number
                master_id_to_file[mid] = mid;
            }

            //Parse all master files keyed by file number
            var file_to_shapes = new Dictionary<string, Dictionary<string, Shape>>();

            foreach (var item in zf.Entries)
            {
                string name = item.Key;

                if (!(name.StartsWith("visio/masters/master") && name.EndsWith(".xml")))
                {
                    continue;
                }

                if (name.Contains("masters.xml"))
                {
                    continue;
                }

                var master_num = Path.GetFileNameWithoutExtension(name).Replace("master", "");

                root = XDocument.Parse(GetFileContent(zf, name)).Root;

                var shapes_data = new Dictionary<string, Shape>();

                foreach (var shape in root.Descendants().Where(item => item.Name.LocalName == "Shape"))
                {
                    var sd = _parse_single_shape(shape);

                    shapes_data.Add(sd.Id, sd);
                }

                if (shapes_data.Count > 0)
                {
                    file_to_shapes.Add(master_num, shapes_data);
                }
            }

            // Re-key by Master ID using the mapping
            foreach (var item in master_id_to_file)
            {
                var mid = item.Key;
                var fnum = item.Value;

                if (file_to_shapes.ContainsKey(fnum))
                {
                    masters[mid] = file_to_shapes[fnum];
                }
            }

            //For any file not mapped (e.g. missing rels), add by file number as fallback
            var mapped_files = master_id_to_file.Values;

            foreach (var item in file_to_shapes)
            {
                var fnum = item.Key;
                var shapes_data = item.Value;

                if (mapped_files.Contains(fnum) == false)
                {
                    masters[fnum] = shapes_data;
                }
            }

            return masters;
        }

        /// <summary>
        /// Parse a single <Shape> element into a rich dict.
        /// </summary>
        /// <param name="shape_elem"></param>
        /// <returns></returns>
        private static Shape _parse_single_shape(XElement shape_elem)
        {
            var sd = new Shape()
            {
                Id = shape_elem.GetAttributeValue("ID"),
                Name = shape_elem.GetAttributeValue("Name"),
                NameU = shape_elem.GetAttributeValue("NameU"),
                Type = shape_elem.GetAttributeValue("Type") ?? "Shape",
                MasterId = shape_elem.GetAttributeValue("Master"),
                MasterShapeId = shape_elem.GetAttributeValue("MasterShape"),
                Cells = new List<Cell>(),
                Geometries = new List<Section>(),
                TextParts = new List<TextInfo>(),
                CharacterFormats = null,
                ParagraphFormats = null,
                SubShapes = new List<Shape>(),
                Controls = new List<Section>(),       // Row_N -> {X, Y, ...}
                Connections = new List<Section>(),    // IX -> {X, Y, ...}
                User = new Section(),                 // User-defined cells (e.g., msvStructureType)
                ForeignData = null,                   // ForeignData info for embedded images
                HyperLinks = new List<HyperLink>(),   // List of {description, address, sub_address, frame}
                LineStyle = shape_elem.GetAttributeValue("LineStyle"),
                FillStyle = shape_elem.GetAttributeValue("FillStyle"),
                TextStyle = shape_elem.GetAttributeValue("TextStyle")
            };

            //Parse top-level cells
            foreach (var cell in shape_elem.Children("Cell"))
            {
                var n = cell.GetAttributeValue("N");
                var v = cell.GetAttributeValue("V");
                var f = cell.GetAttributeValue("F");

                sd.Cells.Add(new Cell { Name = n, Value = v, Formula = f });
            }

            Func<XElement, Section> getSection = (element) =>
            {
                var name = element.GetAttributeValue("N");

                Section section = new Section() { Name = name };

                foreach (var row in element.Children("Row"))
                {
                    var row_ix = row.GetAttributeValue("IX", "0");
                    var row_name = row.GetAttributeValue("N");
                    var row_type = row.GetAttributeValue("T");
                    var isDelete = row.GetAttributeValue("Del") == "1";

                    Row sectionRow = new Row() { Index = row_ix, Name = row_name, Type = row_type, IsDelete = isDelete, Cells = new List<Cell>() };

                    foreach (var cell in row.Children("Cell"))
                    {
                        sectionRow.Cells.Add(new Cell() { Name = cell.GetAttributeValue("N"), Value = cell.GetAttributeValue("V") });
                    }

                    section.Rows ??= new List<Row>();

                    section.Rows.Add(sectionRow);
                }

                return section;
            };

            //Parse Section elements
            foreach (var section in shape_elem.Children("Section"))
            {
                var sec_name = section.GetAttributeValue("N");

                if (sec_name == "Geometry")
                {
                    var geo = _parse_geometry_section(section);

                    if (geo != null)
                    {
                        sd.Geometries.Add(geo);
                    }
                }
                else if (sec_name == "Controls")
                {
                    sd.Controls.Add(getSection(section));
                }
                else if (sec_name == "User")
                {
                    sd.User = getSection(section);
                }
                else if (sec_name == "Connection")
                {
                    sd.Connections.Add(getSection(section));
                }
                else if (sec_name == "Character")
                {
                    sd.CharacterFormats = getSection(section);
                }
                else if (sec_name == "Paragraph")
                {
                    sd.ParagraphFormats = getSection(section);
                }
                else if (sec_name == "FillGradientDef")
                {
                    //Multi-stop gradient definitions
                    var grad_stops = new List<GradientStop>();

                    foreach (var row in section.Children("Row"))
                    {
                        var row_ix = row.GetAttributeValue("IX", "0");
                        var stop_cells = new List<Cell>();

                        foreach (var cell in row.Children("Cell"))
                        {
                            stop_cells.Add(new Cell() { Name = cell.GetAttributeValue("N"), Value = cell.GetAttributeValue("V") });
                        }

                        var pos = GetCellNumberValue(stop_cells, "GradientStopPosition");
                        var color = GetCellValue(stop_cells, "GradientStopColor");
                        var trans = GetCellNumberValue(stop_cells, "GradientStopTransparency");

                        if (!string.IsNullOrEmpty(color))
                        {
                            grad_stops.Add(new GradientStop() { Position = pos * 100, Color = color, Transparency = trans });
                        }
                    }

                    sd.GradientStops ??= new List<GradientStop>();

                    sd.GradientStops.AddRange(grad_stops);
                }
            }

            //Also parse Geom sections that are direct children (alternative format)
            for (var geom_idx = 0; geom_idx < 20; geom_idx++) //Max 20 geometry sections
            {
                var geom_section = shape_elem.Child("Geom");

                if (geom_section != null)
                {
                    break;
                }
            }

            //Parse text
            var text_elem = shape_elem.Child("Text");

            if (text_elem != null)
            {
                sd.Text = text_elem.Value.Trim();
                sd.TextParts.AddRange(_parse_text_element(text_elem));
            }

            //Parse sub-shapes (for groups)
            var shapes_container = shape_elem.Child("Shapes");

            if (shapes_container != null)
            {
                foreach (var sub_shape in shapes_container.Children("Shape"))
                {
                    sd.SubShapes.Add(_parse_single_shape(sub_shape));
                }
            }

            //Parse ForeignData (embedded images)
            sd.ForeignData = _parse_foreign_data(shape_elem);

            //Parse hyperlinks (cross-page references, external links)
            foreach (var sect in shape_elem.Children("Section"))
            {
                if (sect.GetAttributeValue("N") == "Hyperlink")
                {
                    foreach (var row in sect.Children("Row"))
                    {
                        var link = new HyperLink();

                        foreach (var cell in row.Children("Cell"))
                        {
                            var n = cell.GetAttributeValue("N");
                            var v = cell.GetAttributeValue("V");

                            if (n == "Description")
                            {
                                link.Description = v;
                            }
                            else if (n == "Address")
                            {
                                link.Address = v;
                            }
                            else if (n == "SubAddress")
                            {
                                link.SubAddress = v;
                            }
                            else if (n == "Frame")
                            {
                                link.Frame = v;
                            }

                            sd.HyperLinks.Add(link);
                        }
                    }
                }
            }

            return sd;
        }

        /// <summary>
        /// Parse a Geometry section into a list of geometry rows.
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        private static Section _parse_geometry_section(XElement section)
        {
            var geo = new Section() { Name = "Geometry", Rows = new List<Row>() };

            foreach (var cell in section.Children("Cell"))
            {
                var n = cell.GetAttributeValue("N");
                var v = cell.GetAttributeValue("V", "0");

                if (n == "NoFill" && v == "1")
                {
                    geo.NoFill = true;
                }
                else if (n == "NoLine" && v == "1")
                {
                    geo.NoLine = true;
                }
                else if (n == "NoShow" && v == "1")
                {
                    geo.NoShow = true;
                }
            }

            foreach (var row in section.Children("Row"))
            {
                var row_type = row.GetAttributeValue("T");
                var row_ix = row.GetAttributeValue("IX");
                var isDelete = row.GetAttributeValue("Del") == "1";

                var row_data = new Row() { Index = row_ix, Type = row_type, IsDelete = isDelete, Cells = new List<Cell>() };

                foreach (var cell in row.Children("Cell"))
                {
                    var n = cell.GetAttributeValue("N");
                    var v = cell.GetAttributeValue("V");
                    var f = cell.GetAttributeValue("F");

                    row_data.Cells.Add(new Cell() { Name = n, Value = v, Formula = f }); // type: ignore[index]
                }

                geo.Rows.Add(row_data); // type: ignore[attr-defined]    
            }

            //Store section IX for merging
            geo.Index = section.GetAttributeValue("IX", "0");

            return geo;
        }

        /// <summary>
        /// Parse a <Text> element into parts with formatting references.
        /// </summary>
        /// <param name="text_elem"></param>
        /// <returns></returns>
        private static List<TextInfo> _parse_text_element(XElement text_elem)
        {
            var parts = new List<TextInfo>();
            var current_cp = "0";
            var current_pp = "0";

            //Process text content with inline elements

            if (!string.IsNullOrEmpty(text_elem.Value))
            {
                parts.Add(new TextInfo() { Text = text_elem.Value, CP = current_cp, PP = current_pp });
            }

            foreach (var child in text_elem.Elements())
            {
                string name = child.Name.ToString();

                string tag = null;

                if (name.Contains("}"))
                {
                    tag = name.Split('}').Last();
                }
                else
                {
                    tag = name;
                }

                if (tag == "cp")
                {
                    current_cp = child.GetAttributeValue("IX", "0");
                }
                else if (tag == "pp")
                {
                    current_pp = child.GetAttributeValue("IX", "0");
                }
                else if (tag == "fld")
                {
                    // Field element — extract text
                    var field_text = child.Value.Trim();

                    if (!string.IsNullOrEmpty(field_text))
                    {
                        parts.Add(new TextInfo { Text = field_text, CP = current_cp, PP = current_pp });
                    }
                }

                if (child.NextNode != null && child.NextNode.NodeType == XmlNodeType.Text)
                {
                    parts.Add(new TextInfo { Text = (child.NextNode as XText).Value, CP = current_cp, PP = current_pp });
                }
            }

            return parts;
        }

        #endregion

        #region Geometry to SVG path conversion

        /// <summary>
        /// Convert a parsed Geometry section to an SVG path 'd' attribute.
        /// </summary>
        /// <param name="geo"></param>
        /// <param name="w">shape width in inches for relative coordinates.</param>
        /// <param name="h">shape height in inches for relative coordinates.</param>
        /// <param name="master_w">if geometry was inherited from a master, these are the master's original dimensions for coordinate scaling.</param>
        /// <param name="master_h"></param>
        /// <returns></returns>
        private static string _geometry_to_path(Section geo, float w, float h, float master_w = 0.0f, float master_h = 0.0f)
        {
            if (geo.NoShow == true)
            {
                return null;
            }

            // Use absolute dimensions for coordinate calculations — 1D connectors
            // can have negative Width/Height (e.g. Height=-0.867 when EndY < BeginY).
            float abs_w = Math.Abs(w) > 1e-10 ? Math.Abs(w) : 0.0f;
            float abs_h = Math.Abs(h) > 1e-10 ? Math.Abs(h) : 0.0f;

            // Compute scale factors if geometry came from a master with different dims
            var abs_mw = Math.Abs(master_w);
            var abs_mh = Math.Abs(master_h);
            var sx = abs_mw > 1e-6 && Math.Abs(abs_mw - abs_w) > 1e-6 ? abs_w / abs_mw : 1.0f;
            var sy = abs_mh > 1e-6 && Math.Abs(abs_mh - abs_h) > 1e-6 ? abs_h / abs_mh : 1.0f;

            var d_parts = new List<string>();
            float cx = 0.0f, cy = 0.0f;  // Current point (inches)

            int _row_idx = 0;
            var rows = geo.Rows;

            foreach (var row in rows)
            {
                string rt = row.Type;
                var cells = row.Cells;

                // Skip geometry rows where all coordinate cells are truly empty
                // (spurious rows from connectors with partial geometry)
                // Note: "0" IS a valid coordinate, so only skip if V is None/empty
                if (rt == "LineTo" || rt == "ArcTo")
                {
                    var _has_any = false;

                    var xy = new string[2] { "X", "Y" };

                    foreach (var _cn in xy)
                    {
                        var _cv = GetCellValue(cells, _cn);

                        if (_cv != null && !string.IsNullOrEmpty(_cv))
                        {
                            _has_any = true;
                            break;
                        }
                    }

                    //Also check if there's a formula (F attribute) — inherited rows
                    if (!_has_any)
                    {
                        foreach (var _cn in xy)
                        {
                            var _cf = GetCellFormular(cells, _cn);

                            if (!string.IsNullOrEmpty(_cf) && _cf != "Inh")
                            {
                                _has_any = true;
                                break;
                            }
                        }
                    }

                    if (!_has_any)
                    {
                        continue;
                    }
                }

                if (rt == "MoveTo")
                {
                    float x = GetCellNumberValue(cells, "X") * sx;
                    float y = GetCellNumberValue(cells, "Y") * sy;

                    d_parts.Add($"M {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");

                    cx = x;
                    cy = y;

                    //Detect oval: MoveTo followed by all ArcTo with nonzero bulge
                    var remaining = rows.Skip(_row_idx + 1);

                    var remaining_types = remaining.Where(item => !string.IsNullOrEmpty(item.Type)).Select(item => item.Type).ToArray();

                    if (remaining_types.Length >= 3 && remaining_types.Take(remaining_types.Length).All(t => t == "ArcTo"))
                    {
                        //Collect all ArcTo endpoints
                        var arc_points = new List<(float, float)>() { (x, y) };

                        foreach (var ar in remaining)
                        {
                            var art = ar.Type; ////to do

                            if (art != "ArcTo")
                            {
                                break;
                            }

                            var ac = ar.Cells;
                            var ax = GetCellNumberValue(ac, "X") * sx;
                            var ay = GetCellNumberValue(ac, "X") * sy;

                            arc_points.Add((ax, ay));
                        }

                        //Check if it closes back to start and has enough arcs for an oval
                        if (arc_points.Count >= 4)
                        {
                            var first = arc_points[0];
                            var last = arc_points.Last();
                            var dist = Math.Pow((Math.Pow((first.Item1 - last.Item1), 2) + Math.Pow((first.Item2 - last.Item2), 2)), 0.5);

                            if (dist < 0.01f)  // Closed shape
                            {
                                float[] all_x = arc_points.Select(item => (float)item.Item1).ToArray();
                                float[] all_y = arc_points.Select(item => (float)item.Item2).ToArray();

                                var ecx = (all_x.Min() + all_x.Max()) / 2.0f * _INCH_TO_PX;
                                var ecy = (abs_h - (all_y.Min() + all_y.Max()) / 2.0f) * _INCH_TO_PX;
                                var erx = (all_x.Max() - all_x.Min()) / 2.0f * _INCH_TO_PX;
                                var ery = (all_y.Max() - all_y.Min()) / 2.0f * _INCH_TO_PX;

                                if (erx > 0.5 && ery > 0.5)
                                {
                                    d_parts.Clear();
                                    d_parts.Add($"M {(ecx - erx).ToFixed(2)} {ecy.ToFixed(2)}");
                                    d_parts.Add($"A {erx:.2f} {ery:.2f} 0 1 0 {ecx + erx:.2f} {ecy:.2f}");
                                    d_parts.Add($"A {erx:.2f} {ery:.2f} 0 1 0 {ecx - erx:.2f} {ecy:.2f}");
                                    d_parts.Add("Z");

                                    break; // skip remaining geometry rows
                                }
                            }
                        }
                    }
                }
                else if (rt == "RelMoveTo")
                {
                    var x = GetCellNumberValue(cells, "X");
                    var y = GetCellNumberValue(cells, "Y");
                    float ax = x * abs_w;
                    float ay = y * abs_h;

                    d_parts.Add($"M {(ax * _INCH_TO_PX).ToFixed(2)} {((abs_h - ay) * _INCH_TO_PX).ToFixed(2)}");

                    cx = ax;
                    cy = ay;
                }
                else if (rt == "LineTo")
                {
                    float x = GetCellNumberValue(cells, "X") * sx;
                    float y = GetCellNumberValue(cells, "Y") * sy;

                    d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");

                    cx = x;
                    cy = y;
                }
                else if (rt == "RelLineTo")
                {
                    var x = GetCellNumberValue(cells, "X");
                    var y = GetCellNumberValue(cells, "Y");
                    var ax = x * abs_w;
                    var ay = y * abs_h;

                    d_parts.Add($"L {ax * _INCH_TO_PX:.2f} {(abs_h - ay) * _INCH_TO_PX:.2f}");

                    cx = ax;
                    cy = ay;
                }
                else if (rt == "ArcTo")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;
                    var a = GetCellNumberValue(cells, "A") * sy;  // bulge scales with Y

                    // A is the bulge/sagitta of the arc
                    _append_arc(d_parts, cx, cy, x, y, a, abs_h);

                    cx = x;
                    cy = y;
                }
                else if (rt == "EllipticalArcTo")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;
                    var a = GetCellNumberValue(cells, "A") * sx; // control point X
                    var b = GetCellNumberValue(cells, "B") * sy;  // control point Y
                    var c_angle = GetCellNumberValue(cells, "C");  // angle of major axis (radians)
                    var d_ratio = GetCellNumberValue(cells, "D");  // ratio major/minor axis

                    _append_elliptical_arc(d_parts, cx, cy, x, y, a, b, d_ratio, c_angle, abs_h);

                    cx = x;
                    cy = y;
                }
                else if (rt == "RelEllipticalArcTo")
                {
                    // Same as EllipticalArcTo but with relative coordinates (0-1)
                    var x = GetCellNumberValue(cells, "X") * abs_w;
                    var y = GetCellNumberValue(cells, "Y") * abs_h;
                    var a = GetCellNumberValue(cells, "A") * abs_w;  // control point X
                    var b = GetCellNumberValue(cells, "B") * abs_h;  // control point Y
                    var c_angle = GetCellNumberValue(cells, "C");  // angle (radians)
                    var d_ratio = GetCellNumberValue(cells, "D");  // ratio major/minor

                    _append_elliptical_arc(d_parts, cx, cy, x, y, a, b, d_ratio, c_angle, abs_h);

                    cx = x;
                    cy = y;
                }
                else if (rt == "NURBSTo")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;

                    // Parse NURBS formula from E cell for control points
                    var e_val = GetCellValue(cells, "E");

                    List<(float, float)> nurbs_pts = _parse_nurbs_formula(e_val, cx, cy, x, y, sx, sy);

                    if (nurbs_pts != null && nurbs_pts.Count >= 2)
                    {
                        //Use quadratic or cubic Bézier approximation
                        if (nurbs_pts.Count == 2)
                        {
                            // Two control points → cubic Bézier
                            var (cp1x, cp1y) = nurbs_pts[0];
                            var (cp2x, cp2y) = nurbs_pts[1];

                            d_parts.Add($"C {(cp1x * _INCH_TO_PX).ToFixed(2)} {((abs_h - cp1y) * _INCH_TO_PX).ToFixed(2)} {(cp2x * _INCH_TO_PX).ToFixed(2)} {((abs_h - cp2y) * _INCH_TO_PX).ToFixed(2)} {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");
                        }
                        else
                        {
                            //Multiple points → approximate with lines through them
                            foreach (var item in nurbs_pts)
                            {
                                var px_val = item.Item1;
                                var py_val = item.Item2;

                                d_parts.Add($"L {(px_val * _INCH_TO_PX).ToFixed(2)} {((abs_h - py_val) * _INCH_TO_PX).ToFixed(2)}");
                            }

                            d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");
                        }
                    }
                    else
                    {
                        d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");
                    }

                    cx = x;
                    cy = y;
                }
                else if (rt == "RelCurveTo" || rt == "RelCubBezTo")
                {
                    var x = GetCellNumberValue(cells, "X");
                    var y = GetCellNumberValue(cells, "Y");
                    var a = GetCellNumberValue(cells, "A");
                    var b = GetCellNumberValue(cells, "B");
                    var c = GetCellNumberValue(cells, "C");
                    var dd = GetCellNumberValue(cells, "D");

                    //Cubic bezier with relative coordinates
                    var cp1x = a * abs_w;
                    var cp1y = b * abs_h;
                    var cp2x = c * abs_w;
                    var cp2y = dd * abs_h;
                    var ex = x * abs_w;
                    var ey = y * abs_h;

                    d_parts.Add($"C {(cp1x * _INCH_TO_PX).ToFixed(2)} {((abs_h - cp1y) * _INCH_TO_PX).ToFixed(2)} {(cp2x * _INCH_TO_PX).ToFixed(2)} {((abs_h - cp2y) * _INCH_TO_PX).ToFixed(2)} {(ex * _INCH_TO_PX).ToFixed(2)} {((abs_h - ey) * _INCH_TO_PX).ToFixed(2)}");

                    cx = ex;
                    cy = ey;
                }
                else if (rt == "Ellipse")
                {
                    //Full ellipse: center (X,Y), point on major axis (A,B), point on minor axis (C,D)
                    var ex = GetCellNumberValue(cells, "X") * sx;
                    var ey = GetCellNumberValue(cells, "Y") * sy;
                    var ea = GetCellNumberValue(cells, "A") * sx;
                    var eb = GetCellNumberValue(cells, "B") * sy;
                    var ec = GetCellNumberValue(cells, "C") * sx;
                    var ed = GetCellNumberValue(cells, "D") * sy;
                    var rx = Math.Sqrt(Math.Pow((ea - ex), 2) + Math.Pow((eb - ey), 2));
                    var ry = Math.Sqrt(Math.Pow((ec - ex), 2) + Math.Pow((ed - ey), 2));

                    if (rx < 0.001)
                    {
                        rx = 0.001;
                    }

                    if (ry < 0.001)
                    {
                        ry = 0.001;
                    }

                    var cpx = ex * _INCH_TO_PX;
                    var cpy = (abs_h - ey) * _INCH_TO_PX;
                    var rpx = rx * _INCH_TO_PX;
                    var rpy = ry * _INCH_TO_PX;

                    //SVG ellipse as two arcs
                    d_parts.Add($"M {(cpx - rpx).ToFixed(2)} {cpy.ToFixed(2)} A {rpx.ToFixed(2)} {rpy.ToFixed(2)} 0 1 0 {(cpx + rpx).ToFixed(2)} {cpy.ToFixed(2)} A {rpx.ToFixed(2)} {rpy.ToFixed(2)} 0 1 0 {(cpx - rpx).ToFixed(2)} {cpy.ToFixed(2)} Z");
                }
                else if (rt == "PolylineTo")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;

                    //Try to parse the formula for intermediate points
                    var a_cell = cells.FirstOrDefault(item => item.Name == "A");
                    var formula = a_cell?.Formula;
                    var pts = _parse_polyline_formula(formula, abs_w, abs_h);

                    if (a_cell != null && (pts == null || pts.Count == 0))
                    {
                        //Try V cell: semicolon - separated "x,y" pairs from VSD parser
                        var v_val = a_cell.Value;

                        if (v_val.Contains(";") || v_val.Contains(","))
                        {
                            pts = new List<(float, float)>();

                            foreach (var pair in v_val.Split(';'))
                            {
                                var parts = pair.Trim().Split(',');

                                if (parts.Length >= 2)
                                {
                                    var firstItem = parts[0].Split('(').LastOrDefault();

                                    pts.Add((Convert.ToSingle(firstItem) * sx, Convert.ToSingle(parts[1]) * sy));
                                }
                            }
                        }
                    }

                    if (pts != null && pts.Count > 0)
                    {
                        foreach (var item in pts)
                        {
                            var px_val = item.Item1;
                            var py_val = item.Item2;

                            d_parts.Add($"L {(px_val * _INCH_TO_PX).ToFixed(2)} {((abs_h - py_val) * _INCH_TO_PX).ToFixed(2)}");
                        }
                    }

                    d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");

                    cx = x;
                    cy = y;
                }
                else if (rt == "SplineStart")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;

                    d_parts.Add($"M {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");
                    cx = x;
                    cy = y;
                }
                else if (rt == "SplineKnot")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;

                    d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");

                    cx = x;
                    cy = y;
                }
                else if (rt == "InfiniteLine")
                {
                    var x = GetCellNumberValue(cells, "X") * sx;
                    var y = GetCellNumberValue(cells, "Y") * sy;
                    var a = GetCellNumberValue(cells, "A") * sx;
                    var b = GetCellNumberValue(cells, "B") * sy;

                    d_parts.Add($"M {(x * _INCH_TO_PX).ToFixed(2)} {((abs_h - y) * _INCH_TO_PX).ToFixed(2)}");
                    d_parts.Add($"L {(a * _INCH_TO_PX).ToFixed(2)} {((abs_h - b) * _INCH_TO_PX).ToFixed(2)}");

                    cx = a;
                    cy = b;
                }

                _row_idx++;
            }

            var result = string.Join(" ", d_parts);

            //Ensure path starts with M (MoveTo) — invalid paths crash renderers
            if (result != null && !result.StartsWith("M"))
            {
                result = $"M 0.00 0.00 {result}";
            }

            //Auto-close path if the last point is very close to the first MoveTo point
            //This ensures proper fill rendering for closed shapes
            if (result != null && !result.Contains("Z") && d_parts.Count >= 3)
            {
                var first_m = Regex.Match(result, @"M\s+([-+]?[\d.]+)\s+([-+]?[\d.]+)");

                if (first_m.Success)
                {
                    var last_part = d_parts.LastOrDefault();

                    var last_coords = Regex.Matches(last_part, @"([-+]?[\d.]+)");

                    if (last_coords.Count >= 2)
                    {
                        var sx_f = float.Parse(first_m.Groups[1].Value);
                        var sy_f = float.Parse(first_m.Groups[2].Value);

                        var ex_f = float.Parse(last_coords[last_coords.Count - 2].Value);
                        var ey_f = float.Parse(last_coords.Last().Value);

                        if (Math.Abs(sx_f - ex_f) < 0.5 && Math.Abs(sy_f - ey_f) < 0.5)
                        {
                            result += " Z";
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Append an arc segment (ArcTo) using SVG arc command.
        /// </summary>
        /// <param name="d_parts"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="bulge">bulge (A) is the sagitta — distance from the midpoint of the chord to the arc.
        /// If bulge is 0, it's a straight line.
        /// </param>
        /// <param name="h"></param>
        private static void _append_arc(List<string> d_parts, float cx, float cy, float x, float y, float bulge, float h)
        {
            if (Math.Abs(bulge) < 1e-6)
            {
                d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((h - y) * _INCH_TO_PX).ToFixed(2)}");

                return;
            }

            // Compute arc from chord and sagitta
            var dx = x - cx;
            var dy = y - cy;
            var chord = Math.Sqrt(dx * dx + dy * dy);

            if (chord < 1e-10)
            {
                return;
            }

            //Radius from sagitta: r = (chord²/4 + sagitta²) / (2 * |sagitta|)
            var sagitta = Math.Abs(bulge);
            var radius = (chord * chord / 4 + sagitta * sagitta) / (2 * sagitta);

            //Clamp radius to max 5x chord length to prevent absurdly large arcs
            var max_radius = chord * 5.0;

            if (radius > max_radius)
            {
                radius = max_radius;
            }

            var radius_px = radius * _INCH_TO_PX;

            // Determine sweep direction
            var large_arc = sagitta > chord / 2.0 ? 1 : 0;
            var sweep = bulge > 0 ? 0 : 1;

            d_parts.Add($"A {radius_px.ToFixed(2)} {radius_px.ToFixed(2)} 0 {large_arc} {sweep} {(x * _INCH_TO_PX).ToFixed(2)} {((h - y) * _INCH_TO_PX).ToFixed(2)}");
        }

        /// <summary>
        /// Append an elliptical arc segment (EllipticalArcTo).
        /// 
        /// (a,b) = control point, d_ratio = aspect ratio (D cell),
        /// c_angle = rotation angle(C cell).
        /// Approximate with SVG arc.
        /// </summary>
        /// <param name="d_parts"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="a"></param>
        /// <param name="b"></param>
        /// <param name="d_ratio"></param>
        /// <param name="c_angle"></param>
        /// <param name="h"></param>
        private static void _append_elliptical_arc(List<string> d_parts, float cx, float cy, float x, float y, float a, float b, float d_ratio, float c_angle, float h)
        {
            //Compute approximate radius from control point
            var mid_x = (cx + x) / 2.0;
            var mid_y = (cy + y) / 2.0;
            var dist_to_control = Math.Sqrt(Math.Pow((a - mid_x), 2) + Math.Pow((b - mid_y), 2));
            var chord = Math.Sqrt(Math.Pow((x - cx), 2) + Math.Pow((y - cy), 2));

            if (chord < 1e-10)
                return;

            var sagitta = dist_to_control;

            if (sagitta < 1e-6)
            {
                d_parts.Add($"L {(x * _INCH_TO_PX).ToFixed(2)} {((h - y) * _INCH_TO_PX).ToFixed(2)}");
                return;
            }

            var rx = (chord * chord / 4 + sagitta * sagitta) / (2 * sagitta);
            var ry = d_ratio > 0.001 ? rx / d_ratio : rx;  // d_ratio is major/minor ratio
            var angle_deg = c_angle > 0 ? c_angle * 180.0f / Math.PI : 0;

            var rx_px = Math.Abs(rx * _INCH_TO_PX);
            var ry_px = Math.Abs(ry * _INCH_TO_PX);

            if (rx_px < 0.1)
                rx_px = 0.1;

            if (ry_px < 0.1)
                ry_px = 0.1;

            //Determine arc direction from control point position relative to chord.
            var cross = (x - cx) * (b - cy) - (y - cy) * (a - cx);
            var sweep = cross < 0 ? 0 : 1;
            var large_arc = 0;

            d_parts.Add($"A {rx_px.ToFixed(2)} {ry_px.ToFixed(2)} {angle_deg.ToFixed(1)} {large_arc} {sweep} {(x * _INCH_TO_PX).ToFixed(2)} {((h - y) * _INCH_TO_PX).ToFixed(2)}");
        }

        /// <summary>
        /// Evaluate a NURBS curve using De Boor's algorithm.
        /// </summary>
        /// <param name="ctrl_pts">list of (x, y, weight) control points</param>
        /// <param name="knots">knot vector</param>
        /// <param name="degree">curve degree (typically 3)</param>
        /// <param name="num_samples">number of output points for tessellation</param>
        /// <returns>list of (x, y) points along the curve</returns>

        /// <summary>
        /// Parse a NURBS formula and return intermediate control points.
        /// 
        /// NURBS format: NURBS(knotLast, degree, xType, yType, x1,y1,k1,w1, x2,y2,k2,w2, ...)
        /// For degree = 3(cubic), we extract control points and scale them.
        /// </summary>
        /// <param name="e_val"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="ex"></param>
        /// <param name="ey"></param>
        /// <param name="sx"></param>
        /// <param name="sy"></param>
        /// <returns>List of (x, y) control points in shape coordinates.</returns>
        private static List<(float, float)> _parse_nurbs_formula(string e_val, float cx, float cy, float ex, float ey, float sx = 1.0f, float sy = 1.0f)
        {
            if (e_val == null)
            {
                return new List<(float, float)>();
            }

            var m = Regex.Match(e_val, @"NURBS\s*\((.*)\)", RegexOptions.IgnoreCase);

            if (!m.Success)
            {
                return new List<(float, float)>();
            }

            var vals = m.Groups[1].Value.Split(",").Select(item => float.Parse(item.Trim())).ToList();

            if (vals.Count < 8)
            {
                return new List<(float, float)>();
            }

            var knot_last = vals[0];
            var degree = (int)(vals[1]);
            var x_type = (int)(vals[2]);  // 0 = fraction of Width, 1 = absolute
            var y_type = (int)(vals[3]);  // 0 = fraction of Height, 1 = absolute

            // Extract control points (groups of 4: x, y, knot, weight)
            var points = new List<(float, float)>();

            for (var i = 4; i < vals.Count - 3; i += 4)
            {
                var px = vals[i];
                var py = vals[i + 1];

                // knot = vals[i + 2], weight = vals[i + 3] — not used for simple approx
                // When x_type/y_type=0, coords are fractions (0-1) — leave as-is
                // When =1, coords are absolute inches — apply scaling
                if (x_type == 1)
                    px *= sx;
                if (y_type == 1)
                    py *= sy;

                points.Add((px, py));
            }

            // Map NURBS control points to absolute coordinates.
            // When type=0 (fractional), the control points are fractions of the
            // overall curve parameter space from start to end.
            var result = new List<(float, float)>();

            foreach (var item in points)
            {
                var px = item.Item1;
                var py = item.Item2;

                var bx = 0.0f;
                var by = 0.0f;

                if (x_type == 0)
                    bx = cx + px * (ex - cx);
                else
                    bx = px;  // already in absolute shape coords

                if (y_type == 0)
                    by = cy + py * (ey - cy);
                else
                    by = py; // already in absolute shape coords

                result.Add((bx, by));
            }

            return result;
        }

        /// <summary>
        /// Parse a POLYLINE formula to extract points.
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="w"></param>
        /// <param name="h"></param>
        /// <returns></returns>
        private static List<(float, float)> _parse_polyline_formula(string formula, float w, float h)
        {
            //Format: POLYLINE(0, 0, x1, y1, x2, y2, ...)
            var pts = new List<(float, float)>();

            if (string.IsNullOrEmpty(formula))
            {
                return pts;
            }

            var m = Regex.Match(formula, @"POLYLINE\s*\((.*)\)", RegexOptions.IgnoreCase);

            if (!m.Success)
            {
                return pts;
            }

            var vals = m.Groups[1].Value.Split(",").Where(item => float.TryParse(item, out _)).Select(item => float.Parse(item.Trim())).ToList();

            // Skip first two values (flags), then pairs
            for (var i = 2; i < vals.Count - 1; i += 2)
            {
                pts.Add((vals[i], vals[i + 1]));
            }

            return pts;
        }

        #endregion

        #region Shape merging (master inheritance)

        /// <summary>
        /// erge a shape with its master, local values override master values.
        /// 
        /// For sub-shapes in groups, parent_master_id is the group's Master ID,
        /// and the sub - shape's master_shape references a shape within that master.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="masters"></param>
        /// <param name="parent_master_id"></param>
        /// <returns></returns>
        private static Shape _merge_shape_with_master(Shape shape, Dictionary<string, Dictionary<string, Shape>> masters,
                              string parent_master_id = null)
        {
            var master_id = shape.MasterId ?? parent_master_id;
            var master_shape_id = shape.MasterShapeId;

            if (string.IsNullOrEmpty(master_id) || !masters.ContainsKey(master_id))
                return shape;

            var master_shapes = masters[master_id];

            // Find the right master shape
            Shape master_sd = null;

            if (!string.IsNullOrEmpty(master_shape_id) && master_shapes.ContainsKey(master_shape_id))
            {
                master_sd = master_shapes[master_shape_id];
            }
            else if (master_shapes.Count > 0)
            {
                // Use first shape in master
                master_sd = master_shapes.Values.FirstOrDefault();
            }

            if (master_sd == null)
                return shape;

            shape.MasterShape = master_sd;

            // Merge cells: master provides defaults, local overrides.
            // Keep local cells that have either a value (V) or a formula (F),
            // since F="Inh" with V="" means "inherit from master" while F="" with
            // a concrete V means "override master".
            var merged_cells = master_sd.Cells.ToList();

            foreach (var item in shape.Cells)
            {
                if (!string.IsNullOrEmpty(item.Value) || !string.IsNullOrEmpty(item.Formula))
                {
                    string name = item.Name;

                    var index = merged_cells.FindIndex(item => item.Name == name);

                    if (index >= 0)
                    {
                        merged_cells[index] = ObjectHelper.CloneObject<Cell>(item);
                    }
                    else
                    {
                        merged_cells.Add(ObjectHelper.CloneObject<Cell>(item));
                    }
                }
            }

            shape.Cells = merged_cells;

            // Merge geometry: use local if present, otherwise master.
            // If local geometry has fewer rows than master (partial override with F='Inh'),
            // merge row-by-row using IX as key.
            var master_geos = master_sd.Geometries;

            Func<(float? Width, float? Height)> getMasterSize = () =>
            {
                // Store master's original dimensions for geometry coordinate scaling
                var master_w_val = GetCellValue(master_sd.Cells, "Width");
                var master_h_val = GetCellValue(master_sd.Cells, "Height");

                float? width = null;
                float? height = null;

                if (!string.IsNullOrEmpty(master_w_val))
                    width = Convert.ToSingle(master_w_val);

                if (!string.IsNullOrEmpty(master_h_val))
                    height = Convert.ToSingle(master_h_val);

                return (width, height);
            };

            if (!shape.HasGeometry && master_sd.HasGeometry)
            {
                shape.Geometries = master_geos;

                var masterSize = getMasterSize();

                if (masterSize.Width.HasValue)
                    shape.MasterWidth = masterSize.Width.Value;

                if (masterSize.Height.HasValue)
                    shape.MasterHeight = masterSize.Height.Value;
            }
            else if (shape.HasGeometry && master_sd.HasGeometry)
            {
                // Mark that this shape had its own geometry (important for 1D connectors)
                shape.HasOwnGeometry = true;

                // Check if this is a 1D connector -- connectors use their own geometry
                // directly (routed paths), don't merge row-by-row with master.
                var is_1d_shape = Convert.ToBoolean(
                    GetCellNumberValue(shape.Cells, "BeginX") != 0.0f
                     && GetCellNumberValue(shape.Cells, "EndX") != 0.0f
                ) || GetCellValue(shape.Cells, "ObjType") == "2";

                if (!is_1d_shape)
                {
                    // IX-based geometry section merge: local shape may override only
                    // specific geometry sections (by IX).  Missing sections come from
                    // the master, preserving the full shape geometry.
                    var local_by_section_ix = new Dictionary<string, Section>();

                    for (var i = 0; i < shape.Geometries.Count; i++)
                    {
                        var g = shape.Geometries[i];

                        local_by_section_ix.Add(g.Index, g);
                    }

                    var merged_geos = new List<Section>();
                    var master_ixs_seen = new HashSet<string>();

                    // Track whether the merged result uses master-space coordinates
                    // (needing scaling) or instance-space coordinates (no scaling).
                    var needs_master_scaling = false;

                    for (var mi = 0; mi < master_geos.Count; mi++)
                    {
                        var master_geo = master_geos[mi];
                        var mix = master_geo.Index ?? mi.ToString();

                        master_ixs_seen.Add(mix);

                        if (local_by_section_ix.ContainsKey(mix))
                        {
                            var local_geo = local_by_section_ix[mix];
                            var local_rows = local_geo.Rows;
                            var master_rows = master_geo.Rows;

                            // Build IX->row map for local row overrides
                            var local_rows_by_ix = new Dictionary<string, Row>();

                            foreach (var r in local_rows)
                            {
                                var rix = r.Index;

                                if (!string.IsNullOrEmpty(rix))
                                {
                                    local_rows_by_ix.Add(rix, r);
                                }
                            }

                            if (local_rows_by_ix.Count > 0 && local_rows.Count < master_rows.Count)
                            {
                                // Partial row override -- merge master rows with local overrides
                                var merged_rows = new List<Row>();
                                var _has_master_only_rows = false;

                                foreach (var mr in master_rows)
                                {
                                    var mrix = mr.Index;

                                    if (!string.IsNullOrEmpty(mrix) && local_rows_by_ix.ContainsKey(mrix))
                                    {
                                        var lr = local_rows_by_ix[mrix];
                                        var merged_cells_r = mr.Cells.ToList();

                                        foreach (var cv in lr.Cells)
                                        {
                                            if (!string.IsNullOrEmpty(cv.Value))
                                            {
                                                string cn = cv.Name;

                                                var index = merged_cells_r.FindIndex(item => item.Name == cn);

                                                if (index >= 0)
                                                {
                                                    merged_cells_r[index] = ObjectHelper.CloneObject<Cell>(cv);
                                                }
                                                else
                                                {
                                                    merged_cells_r.Add(ObjectHelper.CloneObject<Cell>(cv));
                                                }
                                            }
                                        }

                                        var merged_row = new Row()
                                        {
                                            Type = lr.Type ?? mr.Type,
                                            Cells = merged_cells_r,
                                            Index = mrix
                                        };

                                        merged_rows.Add(merged_row);
                                    }
                                    else
                                    {
                                        merged_rows.Add(mr);
                                        _has_master_only_rows = true;
                                    }

                                    local_geo.Rows = merged_rows;

                                    // Only need master scaling if there are rows that came
                                    // entirely from master (in master coordinate space)

                                    if (_has_master_only_rows)
                                        needs_master_scaling = true;
                                }
                            }
                            else
                            {
                                // Local has same or more rows than master — local
                                // geometry values are already in instance coordinate
                                // space (V values match instance Width/Height), so
                                // no master-to-instance scaling is needed.
                            }

                            // Inherit NoFill/NoLine/NoShow from master if not set locally                            

                            if (!local_geo.NoFill && master_geo.NoFill)
                            {
                                local_geo.NoFill = master_geo.NoFill;
                            }

                            if (!local_geo.NoLine && master_geo.NoLine)
                            {
                                local_geo.NoLine = master_geo.NoLine;
                            }

                            if (!local_geo.NoShow && master_geo.NoShow)
                            {
                                local_geo.NoShow = master_geo.NoShow;
                            }

                            merged_geos.Add(local_geo);
                        }
                        else
                        {
                            // Section not overridden locally -- use master section
                            merged_geos.Add(master_geo);

                            needs_master_scaling = true;
                        }
                    }

                    // Add any local-only sections (IX not in master)
                    foreach (var item in local_by_section_ix)
                    {
                        var lix = item.Key;
                        var lg = item.Value;

                        if (!master_ixs_seen.Contains(lix))
                            merged_geos.Add(lg);
                    }

                    shape.Geometries = merged_geos;

                    // Only store master dims for scaling when the merged geometry
                    // contains rows in master coordinate space (not instance space).
                    if (needs_master_scaling)
                    {
                        var master_w_val = GetCellValue(master_sd.Cells, "Width");
                        var master_h_val = GetCellValue(master_sd.Cells, "Height");

                        if (!string.IsNullOrEmpty(master_w_val))
                            shape.MasterWidth = Convert.ToSingle(master_w_val);
                        if (!string.IsNullOrEmpty(master_h_val))
                            shape.MasterHeight = Convert.ToSingle(master_h_val);
                    }
                }
            }

            #region Makeup x/y
            if (shape.HasGeometry && master_sd.HasGeometry)
            {
                foreach (Section geom in shape.Geometries)
                {
                    var masterGeom = master_sd.Geometries.FirstOrDefault(item => item.Index == geom.Index);

                    if (masterGeom != null)
                    {
                        foreach (var row in geom.Rows)
                        {
                            if (row.IsDelete)
                            {
                                continue;
                            }

                            var xCell = row.Cells.FirstOrDefault(item => item.Name == "X");
                            var yCell = row.Cells.FirstOrDefault(item => item.Name == "Y");

                            var masterGeomRow = masterGeom.Rows.FirstOrDefault(item => item.Index == row.Index);

                            if (masterGeomRow != null)
                            {
                                if (xCell == null)
                                {
                                    var masterXCell = masterGeomRow.Cells.FirstOrDefault(item => item.Name == "X");

                                    if (masterXCell != null)
                                    {
                                        row.Cells.Add(masterXCell);
                                    }
                                }

                                if (yCell == null)
                                {
                                    var masterYCell = masterGeomRow.Cells.FirstOrDefault(item => item.Name == "Y");

                                    if (masterYCell != null)
                                    {
                                        row.Cells.Add(masterYCell);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            // Merge text: use local if present, otherwise master
            if (string.IsNullOrEmpty(shape.Text) && !shape.HasTextElement && !string.IsNullOrEmpty(master_sd.Text) && shape.Type != "Group")
            {
                var txt = master_sd.Text;

                if (!Array.Exists(new string[] { "Label", "Abc", "Table", "Entity", "Class" }, item => item == txt))
                {
                    shape.Text = txt;

                    if (!shape.HasTextElement && master_sd.HasTextElement)
                    {
                        shape.TextParts = master_sd.TextParts;
                    }
                }
            }

            // Merge character and paragraph formats
            if (!shape.HasCharacterFormat && master_sd.HasCharacterFormat)
            {
                shape.CharacterFormats = master_sd.CharacterFormats;

                if (!shape.HasParagraphFormat && master_sd.HasParagraphFormat)
                {
                    shape.ParagraphFormats = master_sd.ParagraphFormats;
                }
            }

            // Merge controls, connections, and user cells
            if (!shape.HasControl && master_sd.HasControl)
            {
                shape.Controls = master_sd.Controls;
            }

            if (!shape.HasConnection && master_sd.HasConnection)
            {
                shape.Connections = master_sd.Connections;
            }

            if (!shape.HasUser && master_sd.HasUser)
            {
                shape.User = master_sd.User;
            }

            // Merge foreign data (embedded images) from master
            if (shape.ForeignData == null && master_sd.ForeignData != null)
            {
                shape.ForeignData = master_sd.ForeignData;
            }

            return shape;
        }

        #endregion

        #region Shape to SVG rendering

        /// <summary>
        /// Get a cell value from a shape.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="name"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private static string _get_cell_val(Shape shape, string name, string defaultValue = null)
        {
            return GetCellValue(shape.Cells, name, defaultValue);
        }

        /// <summary>
        /// Get a cell value as float.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="name"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private static float _get_cell_float(Shape shape, string name, float defaultValue = 0.0f)
        {
            return GetCellNumberValue(shape.Cells, name, defaultValue);
        }

        /// <summary>
        /// Map QuickStyleFillColor index to a theme color.
        /// 
        /// Visio QuickStyle indices:
        /// 0=dk1, 1=lt1, 2=dk2, 3=lt2, 4=accent1, ..., 9=accent6
        /// 100=dk1, 101=lt1, 102=dk2(tinted), 103-108=accent1-6(tinted)
        /// </summary>
        /// <param name="qs_fill_color"></param>
        /// <param name="theme_colors"></param>
        /// <returns></returns>
        private static string _resolve_quickstyle_color(int qs_fill_color, Dictionary<string, string> theme_colors)
        {
            var _qs_map = new Dictionary<int, string>()
            {
                { 0, "dk1" }, {1, "lt1" }, {2, "dk2" }, {3, "lt2" },
                { 4, "accent1" }, {5, "accent2" }, {6, "accent3" },
                { 7, "accent4" }, {8, "accent5" }, {9, "accent6" },
                { 100, "dk1" }, {101, "lt1" }, {102, "dk2" },
                { 103, "accent1" }, {104, "accent2" }, {105, "accent3" },
                { 106, "accent4" }, {107, "accent5" }, {108, "accent6" }
            };

            var name = _qs_map.ContainsKey(qs_fill_color) ? _qs_map[qs_fill_color] : null;

            if (!string.IsNullOrEmpty(name) && theme_colors.ContainsKey(name))
                return theme_colors[name];

            // Default to accent1 for unknown values
            return theme_colors.ContainsKey("accent1") ? theme_colors["accent1"] : null;
        }

        /// <summary>
        /// Compute SVG transform for a shape.
        /// 
        /// Handles PinX/PinY positioning, LocPinX/LocPinY, rotation, and flipping.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="page_h"></param>
        /// <returns>SVG transform attribute value.</returns>
        private static string _compute_transform(Shape shape, float page_h)
        {
            var pin_x = _get_cell_float(shape, "PinX") * _INCH_TO_PX;
            var pin_y = (page_h - _get_cell_float(shape, "PinY")) * _INCH_TO_PX;
            var w = _get_cell_float(shape, "Width");
            var h = _get_cell_float(shape, "Height");

            // Default LocPinX/Y to center of shape if not specified (Visio default)
            var _lpx_val = _get_cell_val(shape, "LocPinX");
            var loc_pin_x = (!string.IsNullOrEmpty(_lpx_val) ? Convert.ToSingle(_lpx_val) : Math.Abs(w) * 0.5) * _INCH_TO_PX;
            var _lpy_val = _get_cell_val(shape, "LocPinY");
            var loc_pin_y_raw = !string.IsNullOrEmpty(_lpy_val) ? Convert.ToSingle(_lpy_val) : Math.Abs(h) * 0.5;
            var loc_pin_y = (Math.Abs(h) - loc_pin_y_raw) * _INCH_TO_PX;  // Flip Y for local pin

            var angle = _get_cell_float(shape, "Angle");
            var flip_x = _get_cell_val(shape, "FlipX") == "1";
            var flip_y = _get_cell_val(shape, "FlipY") == "1";

            var parts = new List<string>();

            //Translate so pin point is at correct page position
            var tx = pin_x - loc_pin_x;
            var ty = pin_y - loc_pin_y;

            parts.Add($"translate({tx.ToFixed(2)},{ty.ToFixed(2)})");

            // Apply rotation around local pin
            if (Math.Abs(angle) > 1e-6)
            {
                var angle_deg = -angle * 180.0 / Math.PI; // Visio angles are CCW, SVG CW

                parts.Add($"rotate({angle_deg.ToFixed(2)},{loc_pin_x.ToFixed(2)},{loc_pin_y.ToFixed(2)})");
            }

            // Apply flips around local pin
            if (flip_x || flip_y)
            {
                var sx = flip_x ? -1 : 1;
                var sy = flip_y ? -1 : 1;

                // Translate to origin, scale, translate back
                parts.Add($"translate({loc_pin_x.ToFixed(2)},{loc_pin_y.ToFixed(2)})");
                parts.Add($"scale({sx},{sy})");
                parts.Add($"translate({(-loc_pin_x).ToFixed(2)},{(-loc_pin_y).ToFixed(2)})");
            }

            return string.Join(" ", parts);
        }

        /// <summary>
        /// valuate simple Visio ShapeSheet formulas.
        /// 
        /// Handles: GUARD(expr), IF(cond,t,f), Width*N, Height*N, simple arithmetic.
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="shape"></param>
        /// <param name="defaultValue"></param>
        /// <returns></returns>
        private static float _eval_simple_formula(string formula, Shape shape, float defaultValue = 0.0f)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return defaultValue;
            }

            var f = formula.Trim();

            while (f.ToUpper().StartsWith("GUARD(") && f.EndsWith(")"))
            {
                f = f.Substring(6, f.Length - 6 - 1).Trim();
            }

            while (f.ToUpper().StartsWith("THEMEGUARD(") && f.EndsWith(")"))
            {
                f = f.Substring(11, f.Length - 11 - 1).Trim();
            }

            var if_match = Regex.Match(f, @"IF\s*\((.+),(.+),(.+)\)", RegexOptions.IgnoreCase);

            if (if_match.Success)
                return _eval_simple_formula(if_match.Groups[2].Value.Trim(), shape, defaultValue);

            var w = _get_cell_float(shape, "Width", 1.0f);
            var h = _get_cell_float(shape, "Height", 1.0f);

            var expr = f;

            expr = Regex.Replace(expr, @"\bWidth\b", w.ToString(), RegexOptions.IgnoreCase);
            expr = Regex.Replace(expr, @"\bHeight\b", h.ToString(), RegexOptions.IgnoreCase);

            if (expr.ToUpper().Contains("THEMEVAL"))
                return defaultValue;

            if (Regex.Match(expr, @"^[\d\s+\-*/.()]+$""").Success)
            {
                Expression exp = new Expression(expr);

                return Convert.ToSingle(exp.Evaluate());
            }

            if (float.TryParse(f, out var val))
            {
                return val;
            }

            return defaultValue;
        }

        /// <summary>
        /// Render a single shape as SVG elements. Returns list of SVG strings.
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="page_h"></param>
        /// <param name="parent_master_id"></param>
        /// <param name="_depth"></param>
        /// <param name="media"></param>
        /// <param name="page_rels"></param>
        /// <param name="used_markers"></param>       
        /// <param name="theme_colors"></param>
        /// <param name="layers"></param>
        /// <param name="gradients"></param>
        /// <param name="has_shadow"></param>
        /// <param name="text_layer"></param>
        /// <returns></returns>
        private static List<string> _render_shape_svg(Shape shape, float page_h,
                      Document document,
                      Dictionary<string, Dictionary<string, Shape>> masters,
                      string parent_master_id = null,
                      int _depth = 0,
                      Dictionary<string, byte[]> media = null,
                      Dictionary<string, string> page_rels = null,
                      HashSet<string> used_markers = null,
                      Dictionary<string, string> theme_colors = null,
                      Dictionary<string, Layer> layers = null,
                      Dictionary<string, Gradient> gradients = null,
                      HashSet<string> has_shadow = null,
                      List<string> text_layer = null)
        {
            shape = _merge_shape_with_master(shape, masters, parent_master_id);

            if (media == null)
                media = new Dictionary<string, byte[]>();
            if (page_rels == null)
                page_rels = new Dictionary<string, string>();
            if (used_markers == null)
                used_markers = new HashSet<string>();
            if (theme_colors == null)
                theme_colors = new Dictionary<string, string>();
            if (layers == null)
                layers = new Dictionary<string, Layer>();
            if (gradients == null)
                gradients = new Dictionary<string, Gradient>();
            if (has_shadow == null)
                has_shadow = new HashSet<string>();

            var lines = new List<string>();

            // text_layer collects text SVG to render on top of all geometry
            // Only collect for top-level shapes (depth 0); sub-shapes render
            // text within their group transform to get correct positioning
            var _collect_text = text_layer != null && _depth == 0;

            // Skip shapes that are invisible or purely connection/control metadata
            var vis_val = _get_cell_val(shape, "Visible");

            if (vis_val == "0")
                return lines;

            // Layer visibility check
            var layer_member = _get_cell_val(shape, "LayerMember");

            if (!string.IsNullOrEmpty(layer_member) && layers != null && layers.Count > 0)
            {
                // LayerMember can be "0", "1", "0;1" etc.
                var layer_ids = layer_member.Split(";").Select(item => item.Trim()).ToList();
                var all_hidden = true;

                foreach (var lid in layer_ids)
                {
                    var layer_info = layers.ContainsKey(lid) ? layers[lid] : null;

                    if (layer_info?.Visible != false)
                    {
                        all_hidden = false;
                        break;
                    }
                }

                if (all_hidden)
                {
                    return lines;
                }
            }

            // Skip shapes with only connection points and no geometry/text (connection markers)
            if (shape.HasConnection
              && !shape.HasGeometry
              && string.IsNullOrEmpty(shape.Text)
              && !shape.HasSubShape)
            {
                return lines;
            }

            // Handle shape type
            var shape_type = shape.Type ?? "Shape";

            var w_inch = _get_cell_float(shape, "Width");
            var h_inch = _get_cell_float(shape, "Height");
            var w_px = Math.Abs(w_inch) * _INCH_TO_PX;
            var h_px = Math.Abs(h_inch) * _INCH_TO_PX;

            // --- Style ---
            var line_weight = _get_cell_float(shape, "LineWeight", 0.01f) * _INCH_TO_PX;

            if (line_weight < 0.5f)
                line_weight = 1.5f;  // Minimum visible stroke width
            else if (line_weight > 20)
                line_weight = 20;

            var line_color = ColorHelper.ResolveColor(_get_cell_val(shape, "LineColor"), theme_colors) ?? "#333333";
            var fill_foregnd = ColorHelper.ResolveColor(_get_cell_val(shape, "FillForegnd"), theme_colors);
            var fill_bkgnd = ColorHelper.ResolveColor(_get_cell_val(shape, "FillBkgnd"), theme_colors);

            // Also try resolving via formula if value is a color index
            var _ff_formula = GetCellFormular(shape.Cells, "FillForegnd");
            var _fb_formula = GetCellFormular(shape.Cells, "FillBkgnd");
            var _lc_formula = GetCellFormular(shape.Cells, "LineColor");

            // Resolve THEMEVAL formulas and QuickStyle colors from theme
            var qs_fill_color_val = _get_cell_val(shape, "QuickStyleFillColor");

            if (theme_colors != null && theme_colors.Count > 0 && !string.IsNullOrEmpty(qs_fill_color_val))
            {
                var qs_fill_color = (int)(Convert.ToSingle(qs_fill_color_val ?? "-1"));
                var _theme_fill = qs_fill_color >= 0 ? _resolve_quickstyle_color(qs_fill_color, theme_colors) : null;

                // When FillForegnd has THEMEVAL("FillColor",...), resolve from theme
                if (!string.IsNullOrEmpty(_ff_formula) && _ff_formula.Contains("THEMEVAL") && _ff_formula.Contains("FillColor"))
                {
                    if (!string.IsNullOrEmpty(_theme_fill))
                    {
                        fill_foregnd = _theme_fill;
                    }
                }

                // When FillBkgnd has THEMEVAL("FillColor2",...), resolve from theme
                if (!string.IsNullOrEmpty(_ff_formula) && _ff_formula.Contains("THEMEVAL") && _ff_formula.Contains("FillColor2"))
                {
                    if (!string.IsNullOrEmpty(_theme_fill))
                    {
                        fill_bkgnd = ColorHelper.LightenColor(_theme_fill, 0.85f);
                    }
                }

                // When FillForegnd is completely absent but QuickStyleFillColor exists,
                // the shape relies entirely on theme for its fill color
                if (string.IsNullOrEmpty(fill_foregnd) && string.IsNullOrEmpty(_ff_formula) && !string.IsNullOrEmpty(_theme_fill) && !ColorHelper.IsBlack(_theme_fill))
                {
                    fill_foregnd = _theme_fill;
                }
            }

            // GUARD(color_index) in Visio stencils are theme accent placeholders.
            // Replace magenta (#FF00FF, color 6) with theme accent or sensible default.
            var _default_accent = "#5B9BD5";  // Visio default accent blue

            Func<string, string> getDefaultThemeColor = (name) =>
            {
                return theme_colors.ContainsKey(name) ? theme_colors[name] : _default_accent;
            };

            if (_ff_formula?.Contains("GUARD") == true && fill_foregnd == "#FF00FF")
            {
                fill_foregnd = getDefaultThemeColor("accent1");
            }

            if (_fb_formula?.Contains("GUARD") == true && fill_bkgnd == "#FF00FF")
            {
                fill_bkgnd = getDefaultThemeColor("accent1");
            }

            // When THEMEVAL formula resolves to black (color index 0) but we have no
            // theme colors, the shape likely wants a theme-derived color, not black.
            // Use Visio's default accent blue as fallback.
            if (!string.IsNullOrEmpty(_ff_formula) && _ff_formula.Contains("THEMEVAL") && ColorHelper.IsBlack(fill_foregnd))
            {
                fill_foregnd = getDefaultThemeColor("accent1");
            }

            if (!string.IsNullOrEmpty(_ff_formula) && _ff_formula.Contains("THEMEVAL") && ColorHelper.IsBlack(fill_bkgnd))
            {
                fill_bkgnd = getDefaultThemeColor("accent1");
            }

            // GUARD(0) in stencils: color index 0 = black, but in stencil context
            // this is a theme placeholder. Replace with accent color.
            if (_ff_formula?.Contains("GUARD") == true && ColorHelper.IsBlack(fill_foregnd))
            {
                fill_foregnd = getDefaultThemeColor("accent1");
            }

            if (_ff_formula?.Contains("GUARD") == true && ColorHelper.IsBlack(fill_bkgnd))
            {
                fill_bkgnd = getDefaultThemeColor("accent1");
            }

            // If fill colors are still empty but formulas reference theme, use accent1
            if (string.IsNullOrEmpty(fill_foregnd) && !string.IsNullOrEmpty(_ff_formula) &&
                (_ff_formula.Contains("THEME") || _ff_formula.Contains("GUARD")))
            {
                fill_foregnd = getDefaultThemeColor("accent1");
            }

            if (string.IsNullOrEmpty(fill_bkgnd) && !string.IsNullOrEmpty(_fb_formula) &&
                        (_fb_formula.Contains("THEME") || _fb_formula.Contains("GUARD")))
            {
                fill_bkgnd = getDefaultThemeColor("accent1");
            }

            // Handle F="Inh" (inherited from theme) — if value couldn't be resolved,
            // use theme colors as fallback
            if (string.IsNullOrEmpty(fill_foregnd) && _ff_formula == "Inh" && theme_colors != null && theme_colors.Count > 0)
            {
                fill_foregnd = getDefaultThemeColor("accent1");
            }

            #region find by FillStyle of master shape
            if (string.IsNullOrEmpty(fill_foregnd) && document != null)
            {
                if (document.HasStyleSheet && document.HasColor)
                {
                    string fillStyleId = shape.MasterShape?.FillStyle;

                    if (!string.IsNullOrEmpty(fillStyleId))
                    {
                        var styleSheet = document.StyleSheets.FirstOrDefault(item => item.Id == fillStyleId);

                        if (styleSheet != null && styleSheet.FillStyle != "0")
                        {
                            styleSheet = document.StyleSheets.FirstOrDefault(item => item.Id == styleSheet.FillStyle);
                        }

                        if (styleSheet != null)
                        {
                            var value = GetCellValue(styleSheet.Cells, "FillForegnd");
                            var formular = GetCellFormular(styleSheet.Cells, "FillForegnd");

                            if (value == "Themed" || formular == "THEMEVAL()")
                            {
                                foreach (var color in document.Colors.Reverse())
                                {
                                    string colorValue = color.Value;

                                    if (theme_colors.ContainsValue(colorValue))
                                    {
                                        fill_foregnd = colorValue;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            } 
            #endregion

            if (string.IsNullOrEmpty(fill_bkgnd) && _fb_formula == "Inh" && theme_colors != null && theme_colors.Count > 0)
            {
                fill_bkgnd = getDefaultThemeColor("accent1");
            }

            //if (!string.IsNullOrEmpty(_lc_formula) && (_lc_formula == "Inh" || _lc_formula.Contains("THEME")))
            //{
            //    if (theme_colors != null && theme_colors.Count > 0)
            //    {
            //        line_color = theme_colors.ContainsKey("dk1") ? theme_colors["dk1"] : line_color;
            //    }
            //}

            // QuickStyleLineColor — resolve line color from theme
            var qs_line_color_val = _get_cell_val(shape, "QuickStyleLineColor");
            if (theme_colors != null && theme_colors.Count > 0 && !string.IsNullOrEmpty(qs_line_color_val))
            {
                var qs_line_color = (int)(Convert.ToSingle(qs_line_color_val ?? "-1"));

                if (qs_line_color >= 0)
                {
                    var _theme_line = _resolve_quickstyle_color(qs_line_color, theme_colors);

                    if (!string.IsNullOrEmpty(_theme_line) && (!string.IsNullOrEmpty(_lc_formula) && _lc_formula.Contains("THEMEVAL")))
                    {
                        line_color = _theme_line;
                    }
                }
            }

            // QuickStyleFontColor — resolve text color from theme  
            var qs_font_color_val = _get_cell_val(shape, "QuickStyleFontColor");

            if (theme_colors != null && theme_colors.Count > 0 && !string.IsNullOrEmpty(qs_font_color_val))
            {
                var qs_font_color = (int)(Convert.ToSingle(qs_font_color_val ?? "-1"));

                if (qs_font_color >= 0)
                {
                    var _theme_font = _resolve_quickstyle_color(qs_font_color, theme_colors);

                    if (!string.IsNullOrEmpty(_theme_font))
                    {
                        // Store for text rendering
                        shape.ThemeTextColor = _theme_font;
                    }
                }
                else if (_lc_formula.Contains("THEMEVAL") && ColorHelper.IsBlack(line_color))
                {
                    // THEMEVAL line color defaulting to black — use dark accent instead
                    line_color = "#1F477D";  // Dark blue, matches Visio default
                }
            }

            var line_pattern = (int)GetCellNumberValue(shape.Cells, "LinePattern", 1.0f); ////to do: V= Themed, F=IF(User.Embellishment=1,1,THEMEVAL())
            var rounding = _get_cell_float(shape, "Rounding") * _INCH_TO_PX;

            // Determine fill
            var fill_pat_int = (int)GetCellNumberValue(shape.Cells, "FillPattern", 1.0f);

            string fill = null;

            if (fill_pat_int == 0)
                fill = "none";
            else if (fill_pat_int == 1)
            {
                // Solid fill
                fill = fill_foregnd ?? fill_bkgnd ?? null;

                if (string.IsNullOrEmpty(fill))
                {
                    // In Visio, solid fill with no explicit color defaults to white
                    // for shapes that belong to a master (stencil shapes).
                    // Standalone shapes without fill colors render as outline-only.
                    if (!string.IsNullOrEmpty(shape.MasterId) || !string.IsNullOrEmpty(shape.MasterShapeId))
                    {
                        fill = "#FFFFFF";
                    }
                    else
                    {
                        fill = "none";
                    }
                }
            }
            else if (fill_pat_int >= 25 && fill_pat_int <= 40)
            {
                // Gradient fill — Visio gradients go from FillBkgnd to FillForegnd
                var start_color = fill_bkgnd ?? "#FFFFFF";
                var end_color = fill_foregnd ?? fill_bkgnd ?? "#CCCCCC";

                // If both colors are the same, use a solid fill instead of gradient
                if (start_color.ToUpper() == end_color.ToUpper())
                    fill = start_color;
                else
                {
                    // If both are the same or start is white, create a visible gradient
                    if (start_color == end_color && end_color != "#FFFFFF")
                    {
                        start_color = ColorHelper.LightenColor(end_color, 0.7f);
                    }

                    var grad_dir = _get_cell_float(shape, "FillGradientDir");

                    // Map Visio gradient direction to angle
                    var grad_angle = 0.0f;
                    if (grad_dir != 0)
                        grad_angle = grad_dir * 45;
                    else
                    {
                        // Map FillPattern to gradient direction when FillGradientDir
                        // is not set (classic Visio gradient patterns)
                        var _pattern_angles = new Dictionary<int, int>()
                        {
                            {25, 0},    // Linear left-to-right
                            {26, 90},   // Linear top-to-bottom
                            {27, 45},   // Diagonal top-left to bottom-right
                            {28, 315},  // Diagonal bottom-left to top-right
                            {29, 0},    // Linear from center (radial approx)
                            {30, 90},   // Vertical from center
                            {33, 0},    // Horizontal
                            {34, 90},   // Vertical
                            {35, 45},   // Diagonal
                            {36, 315},  // Reverse diagonal
                            { 40, 0},    // Horizontal
                        };

                        grad_angle = _pattern_angles.ContainsKey(fill_pat_int) ? _pattern_angles[fill_pat_int] : 0;

                    }

                    var grad_id = $"grad_{shape.Id}_{fill_pat_int}";
                    var is_radial = Array.Exists(new int[] { 29, 30, 31, 32, 37, 38, 39 }, item => item == fill_pat_int);

                    gradients.Add(grad_id, new Gradient()
                    {
                        StartColor = start_color,
                        StopColor = end_color,
                        Angle = grad_angle,
                        IsRadial = is_radial
                    });

                    fill = $"url(#{grad_id})";
                }
            }
            else if (fill_pat_int >= 2)
            {
                // Unknown pattern — approximate with blend
                if (!string.IsNullOrEmpty(fill_bkgnd) && !ColorHelper.IsBlack(fill_bkgnd))
                    fill = fill_bkgnd;
                else if (!string.IsNullOrEmpty(fill_foregnd) && !ColorHelper.IsBlack(fill_foregnd))
                    fill = ColorHelper.LightenColor(fill_foregnd, 0.7f);
                else
                    fill = "none";
            }
            else
            {
                fill = "none";
            }

            // --- Device-type color differentiation ---
            // When shapes have a generic fill (e.g., Visio default blue) and no master
            // stencil or embedded images, infer device type from shape text and apply
            // semantically meaningful fill colors for visual differentiation.
            var _shape_text = shape.Text?.ToLower();
            string stroke = null;

            if (!string.IsNullOrEmpty(fill) && fill.ToUpper() == "#4472C4" && !string.IsNullOrEmpty(_shape_text)
            && string.IsNullOrEmpty(shape.MasterId) && shape.ForeignData == null)
            {
                var _DEVICE_COLORS = new Dictionary<string, string>()
                {
                    {"router", "#2E75B6"},      // Blue
                    {"switch", "#548235"},      // Green
                    {"firewall", "#C00000"},    // Red
                    {"server", "#7030A0"},      // Purple
                    {"workstation", "#BF8F00"}, // Gold
                    {"pc", "#BF8F00"},          // Gold
                    {"laptop", "#BF8F00"},      // Gold
                    {"printer", "#ED7D31"},     // Orange
                    {"phone", "#00B0F0"},       // Light blue
                    {"voip", "#00B0F0"},        // Light blue
                    {"ap", "#00B050"},          // Bright green
                    {"access point", "#00B050"},
                    {"hub", "#808080"},         // Gray
                    {"modem", "#404040"},       // Dark gray
                    {"cloud", "#5B9BD5"},       // Sky blue
                    {"internet", "#5B9BD5"},    // Sky blue
                    {"database", "#7030A0"},    // Purple (like server)
                    {"storage", "#7030A0"},     // Purple
                    {"load", "#FF6600"},        // Orange (load balancer)
                    {"balancer", "#FF6600"},    // Orange
                    {"vpn", "#2F5496"},         // Dark blue
                    {"ids", "#C00000"},         // Red (like firewall)
                    {"ips", "#C00000"},         // Red
                    {"wan", "#002060"},         // Navy
                    {"lan", "#548235"},         // Green
                    {"camera", "#44546A"},      // Slate gray
                    {"desktop", "#BF8F00"},     // Gold (like workstation)
                    {"dhcp", "#2E75B6"},        // Blue (network service)
                    {"dns", "#2E75B6"},         // Blue (network service)
                    {"lb", "#FF6600"},          // Orange (load balancer)
                    {"nas", "#7030A0"},         // Purple (storage)
                    {"ups", "#808080"},         // Gray (infrastructure)
                    {"monitor", "#44546A"},     // Slate gray
                    {"gateway", "#002060"}     // Navy
                };

                var _first_word = _shape_text.Split("-")[0].Split("_")[0].Split()[0].Trim();
                var _device_fill = _DEVICE_COLORS.ContainsKey(_first_word) ? _DEVICE_COLORS[_first_word] : null;

                if (!string.IsNullOrEmpty(_device_fill))
                {
                    // Try two-word match
                    var _two_words = string.Join(" ", _shape_text.Split(' ').Skip(2)).Split("-")[0];
                    _device_fill = _DEVICE_COLORS.ContainsKey(_two_words) ? _DEVICE_COLORS[_two_words] : null;

                    if (!string.IsNullOrEmpty(_device_fill))
                    {
                        fill = _device_fill;

                        // Store on shape so text auto-contrast can see computed fill
                        shape.ComputedFill = _device_fill;

                        // Darken stroke to match device color
                        stroke = ColorHelper.ResolveColor(null, theme_colors) ?? "#333333";
                    }
                }
            }

            // Container detection
            var is_container = false;
            var user_data = shape.User;
            var structure_type = GetStructureType(shape);

            if (structure_type == "Container")
                is_container = true;

            string shape_name = (shape.NameU ?? shape.Name)?.ToLower();

            if (!string.IsNullOrEmpty(shape_name) && (new string[] { "dash square", "container", "swimlane" }).Any(item => shape_name.Contains(item)))
                is_container = true;

            // Shadow support — per-shape shadow with offset, color, transparency
            var shdw_pattern = _get_cell_val(shape, "ShdwPattern");
            var shape_has_shadow = !string.IsNullOrEmpty(shdw_pattern) && shdw_pattern != "0";
            string shadow_attr = null;

            if (shape_has_shadow)
            {
                has_shadow.Add("shadow");
                shadow_attr = " filter=\"url(#shadow)\"";
            }

            // No line if pattern 0
            stroke = line_pattern != 0 ? line_color : "none";
            var stroke_width = line_weight;

            var dash_array = _get_dash_array(line_pattern, stroke_width);

            // Fill opacity from FillForegndTrans (0=opaque, 1=transparent)
            var fill_trans = _get_cell_float(shape, "FillForegndTrans");
            var fill_opacity = 1.0f;

            if (fill_trans > 0 && fill_trans <= 1)
                fill_opacity = 1.0f - fill_trans;
            else if (fill_trans > 1)
                fill_opacity = 1.0f - (fill_trans / 100.0f);  // percentage

            // Container style — ensure visibility but respect actual styles
            if (is_container)
            {
                if (fill == "none" && string.IsNullOrEmpty(fill_foregnd))
                {
                    fill = "#F8F8F8";
                    fill_opacity = 0.5f;
                }
                else if (fill != "none" && fill_opacity > 0.9f)
                {
                    // Containers with opaque fills should be semi-transparent
                    // so contents are visible
                    fill_opacity = Math.Max(0.3f, fill_opacity * 0.5f);
                }

                if (string.IsNullOrEmpty(dash_array) && line_pattern <= 1)
                    dash_array = "8,4";
                if (stroke == "none")
                    stroke = "#AAAAAA";
            }

            // Line transparency
            var line_trans = _get_cell_float(shape, "LineColorTrans");
            var stroke_opacity = 1.0f;
            if (line_trans > 0 && line_trans <= 1)
                stroke_opacity = 1.0f - line_trans;
            else if (line_trans > 1)
                stroke_opacity = 1.0f - (line_trans / 100.0f);

            // Build style string
            var style_parts = new List<string>()
            {
                $"fill=\"{fill}\"",
                $"stroke=\"{stroke}\"",
                $"stroke-width=\"{stroke_width.ToFixed(2)}\""
            };

            if (fill_opacity < 0.99f)
                style_parts.Add($"fill-opacity=\"{fill_opacity.ToFixed(2)}\"");
            if (stroke_opacity < 0.99f)
                style_parts.Add($"stroke-opacity=\"{stroke_opacity.ToFixed(2)}\"");
            if (!string.IsNullOrEmpty(dash_array))
                style_parts.Add($"stroke-dasharray=\"{dash_array}\"");

            var style_str = string.Join(" ", style_parts);

            // --- Check for 1D shape (connector/line) ---
            var begin_y = _get_cell_val(shape, "BeginY");
            var end_y = _get_cell_val(shape, "EndY");

            // --- Check for 1D connector groups ---
            // Some shapes (e.g., BPMN Sequence Flow) are Group type but also 1D connectors
            // Render them as connectors if they have BeginX/EndX
            var begin_x = _get_cell_val(shape, "BeginX");
            var end_x = _get_cell_val(shape, "EndX");
            var is_1d = !string.IsNullOrEmpty(begin_x) && !string.IsNullOrEmpty(end_x);
            var obj_type = _get_cell_val(shape, "ObjType");
            var is_1d_group = (shape_type == "Group" || shape.HasSubShape) && is_1d;

            string transform = null;

            // --- Group shapes ---
            if ((shape_type == "Group" || shape.SubShapes?.Count > 0 == true) && !is_1d_group)
            {
                transform = _compute_transform(shape, page_h);
                var group_master_id = shape.MasterId ?? parent_master_id;

                // Group's local coordinate system uses its own Width x Height
                var group_h = h_inch;

                // If group has no Width/Height, estimate from sub-shapes.
                // Sub-shapes may not yet be merged with master, so also check
                // corresponding master sub-shapes for PinX/PinY/Width/Height.
                if ((Math.Abs(group_h) < 1e-6 || Math.Abs(w_inch) < 1e-6) && shape.HasSubShape)
                {
                    var master_shapes_map = masters != null && !string.IsNullOrEmpty(group_master_id) && masters.ContainsKey(group_master_id) ? masters[group_master_id] : null;
                    var max_sub_x = 0.0f;
                    var max_sub_y = 0.0f;

                    foreach (var sub in shape.SubShapes)
                    {
                        var sub_cells = sub.Cells;

                        // Try page sub-shape cells first, fall back to master
                        var sub_px = GetCellNumberValue(sub_cells, "PinX");
                        var sub_py = GetCellNumberValue(sub_cells, "PinY");
                        var sub_w = Math.Abs(GetCellNumberValue(sub_cells, "Width"));
                        var sub_h = Math.Abs(GetCellNumberValue(sub_cells, "Height"));

                        if (sub_px == 0.0f && sub_py == 0.0f && (sub_cells == null || sub_cells.Count == 0))
                        {
                            // Sub-shape has no cells — look up master sub-shape
                            var ms_id = sub.MasterShapeId;
                            var ms = master_shapes_map.ContainsKey(ms_id) ? master_shapes_map[ms_id] : null;

                            if (ms != null)
                            {
                                var mc = ms.Cells;
                                sub_px = GetCellNumberValue(mc, "PinX");
                                sub_py = GetCellNumberValue(mc, "PinY");
                                sub_w = Math.Abs(GetCellNumberValue(mc, "Width"));
                                sub_h = Math.Abs(GetCellNumberValue(mc, "Height"));
                            }
                        }

                        max_sub_x = Math.Max(max_sub_x, sub_px + sub_w / 2.0f);
                        max_sub_y = Math.Max(max_sub_y, sub_py + sub_h / 2.0f);
                    }

                    if (max_sub_y > 0 && Math.Abs(group_h) < 1e-6)
                    {
                        group_h = max_sub_y;
                        h_inch = group_h;
                        h_px = Math.Abs(group_h) * _INCH_TO_PX;
                    }

                    if (max_sub_x > 0 && Math.Abs(w_inch) < 1e-6)
                    {
                        w_inch = max_sub_x;
                        w_px = Math.Abs(w_inch) * _INCH_TO_PX;
                    }
                }

                // Apply clipping only for large groups (containers/swimlanes),
                // not for small stencil/icon groups where sub-shapes may extend
                // slightly beyond the nominal group bounds.
                var _has_text_subs = shape.SubShapes?.Any(item => !string.IsNullOrEmpty(item.Text)) == true;
                var use_clip = w_px > 300 && h_px > 200 && !_has_text_subs;
                string clip_attr = null;

                if (use_clip)
                {
                    var clip_id = $"clip_{shape.Id}";

                    // Add small padding (5%) to avoid cutting off edges
                    var pad_x = w_px * 0.12f;
                    var pad_y = h_px * 0.12f;

                    lines.Add(
                        $@"<defs><clipPath id=""{clip_id}""><rect x=""{(-pad_x).ToFixed(2)}"" y=""{(-pad_y).ToFixed(2)}"" width=""{(w_px + 2 * pad_x).ToFixed(2)}"" height=""{(h_px + 2 * pad_y).ToFixed(2)}""/></clipPath></defs>");

                    clip_attr = $" clip-path=\"url(#{clip_id})\"";
                }

                lines.Add($"<g transform=\"{transform}\"{clip_attr}{shadow_attr}>");

                // Render the group's own geometry (if any)
                if (shape.HasGeometry)
                {
                    var master_w = shape.MasterWidth ?? 0.0f;
                    var master_h = shape.MasterHeight ?? 0.0f;

                    foreach (var geo in shape.Geometries)
                    {
                        string path_d = _geometry_to_path(geo, w_inch, h_inch, master_w, master_h);

                        if (string.IsNullOrEmpty(path_d))
                            continue;

                        var geo_fill = fill;
                        var geo_stroke = stroke;

                        if (geo.NoFill)
                            geo_fill = "none";

                        // Detect open paths: if start != end, don't fill
                        // (SVG auto-closes filled paths, creating visual artifacts)
                        if (geo_fill != "none" && !path_d.Contains("Z"))
                        {
                            var _coords = Regex.Matches(path_d, @"[-+]?[\d.]+");

                            if (_coords.Count >= 4)
                            {
                                var _sx = float.Parse(_coords[0].Value);
                                var _sy = float.Parse(_coords[1].Value);

                                var _ex = float.Parse(_coords[_coords.Count - 2].Value);
                                var _ey = float.Parse(_coords.Last().Value);

                                if (Math.Abs(_sx - _ex) > 2.0f || Math.Abs(_sy - _ey) > 2.0f)
                                {
                                    geo_fill = "none";
                                }
                            }
                        }

                        if (geo.NoLine)
                            geo_stroke = "none";

                        var geo_style = $"fill=\"{geo_fill}\" stroke=\"{geo_stroke}\" stroke-width=\"{stroke_width.ToFixed(2)}\"";

                        if (fill_opacity < 0.99f && geo_fill != "none")
                            geo_style += $" fill-opacity=\"{fill_opacity.ToFixed(2)}\"";
                        if (stroke_opacity < 0.99f && geo_stroke != "none")
                            geo_style += $" stroke-opacity=\"{stroke_opacity.ToFixed(2)}\"";
                        if (!string.IsNullOrEmpty(dash_array))
                            geo_style += $" stroke-dasharray=\"{dash_array}\"";

                        lines.Add($"<path d=\"{path_d}\" {geo_style}{shadow_attr}/>");
                    }
                }

                // Render embedded image for the group
                var fd = shape.ForeignData;

                if (fd != null && media != null && media.Count > 0)
                {
                    string img_href = null;
                    string target = null;
                    string img_name = null;

                    if (!string.IsNullOrEmpty(fd.RelId) && page_rels.ContainsKey(fd.RelId))
                    {
                        target = page_rels[fd.RelId];
                        img_name = target.Split('/').LastOrDefault();

                        if (media.ContainsKey(img_name))
                            img_href = _image_to_data_uri(media[img_name], img_name);
                    }

                    if (!string.IsNullOrEmpty(img_href))
                    {
                        var img_w_px = w_px;
                        var img_h_px = h_px;

                        // Enforce minimum icon size
                        if (img_w_px > 0 && img_h_px > 0)
                        {
                            if (img_w_px < 24 || img_h_px < 24)
                            {
                                var scale = Math.Max(24.0f / img_w_px, 24.0f / img_h_px);
                                img_w_px *= scale;
                                img_h_px *= scale;
                            }
                        }

                        // Use ImgOffsetX/Y for correct positioning within group
                        var img_off_x = _get_cell_float(shape, "ImgOffsetX") * _INCH_TO_PX;
                        var img_off_y = _get_cell_float(shape, "ImgOffsetY") * _INCH_TO_PX;

                        lines.Add(
                                $@"<image x=""{img_off_x.ToFixed(2)}"" y=""{img_off_y.ToFixed(2)}"" 
                                width=""{img_w_px.ToFixed(2)}"" height=""{img_h_px.ToFixed(2)}"" 
                                xlink:href=""{img_href}"" 
                                preserveAspectRatio=""xMidYMid meet""/>"
                            );
                    }
                }

                foreach (var sub in shape.SubShapes)
                {
                    lines.AddRange(_render_shape_svg(
                    sub, group_h, document, masters, group_master_id, _depth + 1,
                    media, page_rels, used_markers,
                    theme_colors, layers, gradients, has_shadow,
                    text_layer = text_layer));
                }

                lines.Add("</g>");

                // Render text for the group itself (but not auto-generated name labels
                // for groups — sub-shapes already provide visible content)
                if (!string.IsNullOrEmpty(shape.Text))
                {
                    if (_collect_text)
                        _append_text_svg(text_layer, shape, page_h, w_px, h_px, theme_colors); // type: ignore[arg-type]
                    else
                        _append_text_svg(lines, shape, page_h, w_px, h_px, theme_colors);
                }

                return lines;
            }

            // --- Compute transform ---
            transform = _compute_transform(shape, page_h);

            // --- Geometry rendering ---
            var has_geometry = shape.HasGeometry;

            // If ALL geometry sections are NoShow, it's likely a conditional-visibility
            // master (e.g., mind map Topic shapes). Force-show the last section that
            // looks like a basic shape outline (has MoveTo + LineTo rows).
            var geom = shape.Geometries;

            if (has_geometry && geom.All(g => g.NoShow == true))
            {
                // Find the best fallback geometry section
                Section _fallback_geo = null;

                for (var i = geom.Count - 1; i >= 0; i--)
                {
                    var _fg = geom[i];
                    var _row_types = _fg.Rows.Select(item => item.Type).ToArray();

                    if (_row_types.Contains("MoveTo") && (_row_types.Contains("LineTo") || _row_types.Contains("ArcTo")))
                    {
                        _fallback_geo = _fg;
                        break;
                    }
                }

                if (_fallback_geo == null && geom != null && geom.Count > 0)
                {
                    // Use any section with rows
                    for (var i = geom.Count - 1; i >= 0; i--)
                    {
                        var _fg = geom[i];

                        if (_fg.HasRow)
                        {
                            _fallback_geo = _fg;
                            break;
                        }
                    }
                }

                if (_fallback_geo != null)
                {
                    _fallback_geo.NoShow = false;
                }
            }

            // For 1D connectors, use dedicated rendering even if they have master geometry
            var is_connector = is_1d || obj_type == "2";

            if (is_connector && is_1d)
            {
                // Ensure connector lines are visible (minimum 2.0px)
                if (stroke_width < 1.0f)
                    stroke_width = 1.5f;

                // 1D shape — check for geometry (routed connectors) first
                var bx = Convert.ToSingle(begin_x) * _INCH_TO_PX;
                var by = (page_h - Convert.ToSingle(begin_y)) * _INCH_TO_PX;
                var ex_px = Convert.ToSingle(end_x) * _INCH_TO_PX;
                var ey_px = (page_h - Convert.ToSingle(end_y)) * _INCH_TO_PX;

                // Arrow markers
                var begin_arrow = (int)(GetCellNumberValue(shape.Cells, "BeginArrow"));
                var end_arrow = (int)(GetCellNumberValue(shape.Cells, "EndArrow"));

                // Default to an arrow if the shape looks like a connector and has no
                // EndArrow cell at all. Check ObjType=2 or name contains "connector".
                var _is_named_connector = shape_name?.Contains("connector") == true;

                if (end_arrow == 0 && (obj_type == "2" || _is_named_connector))
                {
                    var _ea_cell = shape.Cells.FirstOrDefault(item => item.Name == "EndArrow");

                    if ((_ea_cell == null) || (string.IsNullOrEmpty(_ea_cell.Value) || string.IsNullOrEmpty(_ea_cell.Formula)))
                        end_arrow = 4;
                }

                var begin_arrow_size = (int)(GetCellNumberValue(shape.Cells, "BeginArrowSize", 2.0f));
                var end_arrow_size = (int)(GetCellNumberValue(shape.Cells, "EndArrowSize", 2.0f));
                var marker_color = stroke != "none" ? stroke.TrimStart('#') : "333333";
                string marker_attrs = null;

                if (begin_arrow > 0)
                {
                    var mid = $"arrow_start_{begin_arrow_size}_{marker_color}";

                    used_markers.Add(mid);

                    marker_attrs += $" marker-start=\"url(#{mid})\"";
                }

                if (end_arrow > 0)
                {
                    var mid = $"arrow_end_{end_arrow_size}_{marker_color}";

                    used_markers.Add(mid);

                    marker_attrs += $" marker-end=\"url(#{mid})\"";
                }

                // Convert connector geometry to page-coordinate polyline
                // For connectors without PinX, skip geometry and use BeginX→EndX directly
                if (has_geometry && !HasCell(shape.Cells, "PinX"))
                {
                    has_geometry = false;  // Force straight line fallback
                }

                if (has_geometry)
                {
                    // Build polyline from geometry rows in page coordinates
                    // Transform local geo coords to page space using PinX/PinY
                    var pin_x = _get_cell_float(shape, "PinX");
                    var pin_y = _get_cell_float(shape, "PinY");
                    var loc_pin_x = _get_cell_float(shape, "LocPinX");
                    var loc_pin_y = _get_cell_float(shape, "LocPinY");
                    var angle = _get_cell_float(shape, "Angle");
                    float cos_a = Math.Abs(angle) > 1e-6 ? (float)Math.Cos(angle) : 1.0f;
                    float sin_a = Math.Abs(angle) > 1e-6 ? (float)Math.Sin(angle) : 0.0f;

                    var points = new List<(float, float)>();
                    var has_move_to = false;

                    foreach (var geo in shape.Geometries)
                    {
                        if (geo.NoShow)
                            continue;

                        foreach (var row in geo.Rows)
                        {
                            var rt = row.Type;
                            var cells = row.Cells;

                            if (Array.Exists(new string[] {"MoveTo", "LineTo", "ArcTo",
                              "EllipticalArcTo", "NURBSTo",
                              "SplineStart", "SplineKnot" }, item => item == rt))
                            {
                                var x_val = GetCellValue(cells, "X");
                                var y_val = GetCellValue(cells, "Y");

                                // Skip rows with incomplete coordinates — both X
                                // and Y must be present for a valid connector point
                                if (string.IsNullOrEmpty(x_val) && x_val != "0")
                                    continue;

                                if (string.IsNullOrEmpty(y_val) && y_val != "0")
                                    continue;

                                if (rt == "MoveTo")
                                    has_move_to = true;

                                var lx = float.Parse(x_val);
                                var ly = float.Parse(y_val);

                                // Local to page: translate by pin offset
                                var dx = lx - loc_pin_x;
                                var dy = ly - loc_pin_y;
                                var px = pin_x + dx * cos_a - dy * sin_a;
                                var py = pin_y + dx * sin_a + dy * cos_a;

                                // To SVG pixels
                                var sx_px = px * _INCH_TO_PX;
                                var sy_px = (page_h - py) * _INCH_TO_PX;

                                points.Add((sx_px, sy_px));
                            }
                        }
                    }

                    // If geometry had no MoveTo, insert the begin point as start
                    if (points.Count > 0 && !has_move_to)
                        points.Insert(0, (bx, by));

                    if (points.Count >= 2)
                    {
                        var d_parts = new List<string>() { $"M {points[0].Item1.ToFixed(2)} {points[0].Item2.ToFixed(2)}" };

                        foreach (var pt in points.Skip(1))
                            d_parts.Add($"L {pt.Item1.ToFixed(2)} {pt.Item2.ToFixed(2)}");

                        var path_d = string.Join(" ", d_parts);

                        lines.Add($@"<path d=""{path_d}"" fill=""none"" stroke=""{stroke}"" stroke-width=""{stroke_width.ToFixed(2)}""" + (!string.IsNullOrEmpty(dash_array) ? $@" stroke-dasharray=""{dash_array}""" : "")
                            + marker_attrs
                            + "/>");
                    }
                    else
                    {
                        // Fallback to straight line
                        lines.Add(
                            $@"<line x1=""{bx.ToFixed(2)}"" y1= ""{by.ToFixed(2)}"" x2=""{ex_px.ToFixed(2)}"" y2=""{ey_px.ToFixed(2)}"" stroke=""{stroke}"" stroke-width= ""{stroke_width.ToFixed(2)}"""
                            + (!string.IsNullOrEmpty(dash_array) ? $" stroke-dasharray=\"{dash_array}\"" : "")
                            + marker_attrs
                            + "/>"
                        );
                    }
                }
                else
                {
                    // No geometry — simple straight line
                    lines.Add(
                        $@"<line x1=""{bx.ToFixed(2)}"" y1=""{by.ToFixed(2)}"" x2=""{ex_px.ToFixed(2)}"" y2=""{ey_px.ToFixed(2)}""         
                        stroke=""{stroke}"" stroke-width=""{stroke_width.ToFixed(2)}""" + (!string.IsNullOrEmpty(dash_array) ? $" stroke-dasharray=\"{dash_array}\"" : "")
                        + marker_attrs
                        + "/>"
                        );
                }
            }
            else if (has_geometry)
            {
                // 2D shape with geometry
                var master_w = shape.MasterWidth ?? 0.0f;
                var master_h = shape.MasterHeight ?? 0.0f;

                foreach (var geo in shape.Geometries)
                {
                    var path_d = _geometry_to_path(geo, w_inch, h_inch, master_w, master_h);

                    if (string.IsNullOrEmpty(path_d))
                        continue;

                    var geo_fill = fill;
                    var geo_stroke = stroke;

                    if (geo.NoFill)
                        geo_fill = "none";

                    if (geo.NoLine)
                        geo_stroke = "none";

                    var geo_style = $"fill=\"{geo_fill}\" stroke=\"{geo_stroke}\" stroke-width=\"{stroke_width.ToFixed(2)}\"";

                    if (fill_opacity < 0.99f && geo_fill != "none")
                        geo_style += $" fill-opacity=\"{fill_opacity.ToFixed(2)}\"";
                    if (stroke_opacity < 0.99f && geo_stroke != "none")
                        geo_style += $" stroke-opacity=\"{stroke_opacity.ToFixed(2)}\"";
                    if (!string.IsNullOrEmpty(dash_array))
                        geo_style += $" stroke-dasharray=\"{dash_array}\"";

                    lines.Add($"<path d=\"{path_d}\" {geo_style}{shadow_attr} transform=\"{transform}\"/>");
                }
            }
            else
            {
                // No geometry, no 1D — fall back to outlined rectangle
                if (w_px > 0 && h_px > 0 && (fill != "none" || !string.IsNullOrEmpty(shape.Text)))
                {
                    var rx_val = Math.Max(rounding, 4.0f); // Slightly rounded for aesthetics
                    var rx_attr = $" rx=\"{rx_val.ToFixed(2)}\"";

                    // For fallback shapes, prefer outlined rectangle over filled
                    var fallback_fill = fill != "none" ? fill : "#FAFAFA";
                    var fallback_stroke = stroke != "none" ? stroke : (ColorHelper.ResolveColor(_get_cell_val(shape, "LineColor"), theme_colors) ?? "#CCCCCC");

                    var fallback_style = $"fill=\"{fallback_fill}\" stroke=\"{fallback_stroke}\" stroke-width=\"{Math.Max(stroke_width, 0.75f).ToFixed(2)}\"";

                    if (fill_opacity < 0.99f && fallback_fill != "none")
                        fallback_style += $" fill-opacity=\"{fill_opacity.ToFixed(2)}\"";
                    if (stroke_opacity < 0.99f && fallback_stroke != "none")
                        fallback_style += $" stroke-opacity=\"{stroke_opacity.ToFixed(2)}\"";
                    if (!string.IsNullOrEmpty(dash_array))
                        fallback_style += $" stroke-dasharray=\"{dash_array}\"";

                    lines.Add($"<rect x=\"0\" y=\"0\" width=\"{w_px.ToFixed(2)}\" height=\"{h_px.ToFixed(2)}\" {fallback_style}{rx_attr} transform=\"{transform}\"/>");

                    // If shape has no text but has a name, show name as label
                    // Only for top-level shapes (depth 0), not sub-shapes in stencils
                    if (string.IsNullOrEmpty(shape.Text) && !shape.HasTextElement && _depth == 0)
                    {
                        string shape_label = shape.NameU ?? shape.Name;

                        if (!string.IsNullOrEmpty(shape_label) && !shape_label.StartsWith("Sheet."))
                        {
                            // Clean up generic names
                            shape_label = shape_label.Contains(".") ? shape_label.Split(".").LastOrDefault() : shape_label;

                            if (!string.IsNullOrEmpty(shape_label) && shape_label.Length < 30 && !float.TryParse(shape_label, out _))
                                shape.Text = shape_label;
                        }
                    }
                }
            }

            //--- Embedded image rendering ---
            var fdata = shape.ForeignData;

            if (fdata != null && media != null && media.Count > 0)
            {
                string img_href = null;

                if (!string.IsNullOrEmpty(fdata.RelId) && page_rels.ContainsKey(fdata.RelId))
                {
                    string target = page_rels[fdata.RelId];
                    string img_name = target.Split('/').LastOrDefault();

                    if (media.ContainsKey(img_name))
                    {
                        // Always use data URIs — cairosvg doesn't resolve file paths
                        img_href = _image_to_data_uri(media[img_name], img_name);
                    }
                }
                else if (!string.IsNullOrEmpty(fdata.Data))
                {
                    var ext_map = new Dictionary<string, string>()
                    {
                        { "PNG", ".png" }, {"JPEG", ".jpeg" }, {"BMP", ".bmp" },
                        { "GIF", ".gif" }, {"TIFF", ".tiff" }
                    };

                    string comp = (fdata.CompressionType ?? "PNG").ToUpper();
                    var fake_ext = ext_map.ContainsKey(comp) ? ext_map[comp] : ".png";

                    var raw = Convert.FromBase64String(fdata.Data);
                    var fname = $"inline_{shape.Id}{fake_ext}";
                    img_href = _image_to_data_uri(raw, fname);
                }

                if (!string.IsNullOrEmpty(img_href))
                {
                    var img_w = Math.Max(_get_cell_float(shape, "ImgWidth"), w_inch);
                    var img_h = Math.Max(_get_cell_float(shape, "ImgHeight"), h_inch);
                    var img_off_x = _get_cell_float(shape, "ImgOffsetX");
                    var img_off_y = _get_cell_float(shape, "ImgOffsetY");
                    var img_w_px = img_w * _INCH_TO_PX;
                    var img_h_px = img_h * _INCH_TO_PX;

                    // Enforce minimum icon size of 24x24px
                    if (img_w_px > 0 && img_h_px > 0)
                    {
                        if (img_w_px < 24 || img_h_px < 24)
                        {
                            var scale = Math.Max(24.0f / img_w_px, 24.0f / img_h_px);
                            img_w_px *= scale;
                            img_h_px *= scale;
                        }

                        var img_x_px = img_off_x * _INCH_TO_PX;
                        var img_y_px = img_off_y * _INCH_TO_PX;

                        lines.Add($@"<image x=""{img_x_px.ToFixed(2)}"" y=""{img_y_px.ToFixed(2)}"" width=""{img_w_px.ToFixed(2)}"" height= ""{img_h_px.ToFixed(2)}"" 
                                     xlink:href=""{img_href}"" preserveAspectRatio=""xMidYMid meet"" transform=""{transform}""/>");
                    }
                }
            }

            // --- Text rendering ---
            if (!string.IsNullOrEmpty(shape.Text))
            {
                if (_collect_text)
                    _append_text_svg(text_layer, shape, page_h, w_px, h_px, theme_colors); // type: ignore[arg-type]
                else
                    _append_text_svg(lines, shape, page_h, w_px, h_px, theme_colors);
            }

            // No fallback rectangle for shapes inside groups (sub-shapes)
            // is handled by skipping the else branch when geometry/1D absent
            // and the shape has no meaningful content

            return lines;
        }

        /// <summary>
        /// Map a Visio font name to an SVG-compatible font-family string.
        /// </summary>
        /// <param name="font_name"></param>
        /// <returns></returns>
        private static string _map_font_family(string font_name)
        {
            if (string.IsNullOrEmpty(font_name) || font_name == "Themed")
                return "Noto Sans, sans-serif";

            var key = font_name.ToLower().Trim();

            if (_FONT_MAP.ContainsKey(key))
                return _FONT_MAP[key];

            // Keep original font with fallbacks
            return $"{font_name}, Noto Sans, sans-serif";
        }

        /// <summary>
        /// Append SVG text elements for a shape's text.
        /// 
        ///  Uses clipPath to constrain text within shape bounding box.
        ///  Supports multi-format text(text_parts with different cp indices).
        /// </summary>
        /// <param name="lines"></param>
        /// <param name="shape"></param>
        /// <param name="page_h"></param>
        /// <param name="w_px"></param>
        /// <param name="h_px"></param>
        /// <param name="theme_colors"></param>
        private static void _append_text_svg(List<string> lines, Shape shape, float page_h,
                     float w_px, float h_px, Dictionary<string, string> theme_colors = null)
        {
            string text = shape.Text;

            if (string.IsNullOrEmpty(text))
                return;

            // Text position
            var pin_x = _get_cell_float(shape, "PinX") * _INCH_TO_PX;
            var pin_y = (page_h - _get_cell_float(shape, "PinY")) * _INCH_TO_PX;
            var w_inch = _get_cell_float(shape, "Width");
            var h_inch = _get_cell_float(shape, "Height");
            var loc_pin_x = _get_cell_float(shape, "LocPinX") * _INCH_TO_PX;
            var loc_pin_y = _get_cell_float(shape, "LocPinY") * _INCH_TO_PX;

            // Default LocPin to center of shape when not explicitly set
            if (loc_pin_x == 0 && !HasCell(shape.Cells, "LocPinX"))
            {
                loc_pin_x = Math.Abs(w_inch) * 0.5f * _INCH_TO_PX;
            }

            if (loc_pin_y == 0 && !HasCell(shape.Cells, "LocPinY"))
            {
                loc_pin_y = Math.Abs(h_inch) * 0.5f * _INCH_TO_PX;
            }

            // Detect 1D connector shapes - offset text above the line
            var _is_1d_shape = !string.IsNullOrEmpty(_get_cell_val(shape, "BeginX")) && !string.IsNullOrEmpty(_get_cell_val(shape, "EndX"));

            // Text block offset
            var txt_pin_x = _get_cell_float(shape, "TxtPinX");
            var txt_pin_y = _get_cell_float(shape, "TxtPinY");

            // Calculate text center in page coordinates
            // Use TxtPinX/Y when explicitly set (including negative values for below-shape labels)
            var _has_txt_pin = HasCell(shape.Cells, "TxtPinX") || HasCell(shape.Cells, "TxtPinY");

            float tx = 0.0f;
            float ty = 0.0f;
            float shape_left = 0.0f;

            if (_has_txt_pin || txt_pin_x != 0 || txt_pin_y != 0)
            {
                shape_left = pin_x - loc_pin_x;
                var shape_top = pin_y - (Math.Abs(h_inch) * _INCH_TO_PX - loc_pin_y);
                var _eff_txt_pin_x = txt_pin_x != 0 ? txt_pin_x : Math.Abs(w_inch) * 0.5f;
                var _eff_txt_pin_y = txt_pin_y;
                tx = shape_left + _eff_txt_pin_x * _INCH_TO_PX;
                ty = shape_top + (Math.Abs(h_inch) - _eff_txt_pin_y) * _INCH_TO_PX;
            }
            else
            {
                // Default to shape center
                shape_left = pin_x - loc_pin_x;
                var shape_top = pin_y - (Math.Abs(h_inch) * _INCH_TO_PX - loc_pin_y);
                tx = shape_left + Math.Abs(w_inch) * _INCH_TO_PX * 0.5f;
                ty = shape_top + Math.Abs(h_inch) * _INCH_TO_PX * 0.5f;
            }

            // Get text formatting — support multi-format text via text_parts
            var char_formats = shape.CharacterFormats;
            var char_fmt = char_formats?.Rows?.FirstOrDefault(item => item.Index == "0");
            float font_size = (char_fmt != null ? GetCellNumberValue(char_fmt.Cells, "Size") : 0.1111f) * _INCH_TO_PX;

            // Auto-scale font for small shapes: when the default 8pt font is used
            // and the shape has enough space, increase font to be clearly readable.
            var _font_was_default = (font_size < 8.5 && font_size > 7.5);

            if (font_size < 6)
                font_size = 8;
            else if (font_size > 72)
                font_size = 72;

            if (_font_was_default && w_px > 40 && h_px > 20)
            {
                // Scale font to fit comfortably: ~60% of height, capped by width
                var _text_len = shape.Text?.Length ?? 0;

                if (_text_len > 0)
                {
                    var _max_by_height = h_px * 0.45f;
                    var _max_by_width = w_px * 0.85f / (_text_len * 0.55f);
                    var _auto_size = Math.Min(_max_by_height, _max_by_width);
                    font_size = Math.Max(font_size, Math.Min(_auto_size, 16.0f));
                }
            }

            var text_color = ColorHelper.ResolveColor(GetCellValue(char_fmt?.Cells, "Color"), theme_colors) ?? "#000000";

            // Use theme text color if available and char color is default
            if (text_color == "#000000" && !string.IsNullOrEmpty(shape.ThemeTextColor))
                text_color = shape.ThemeTextColor;

            // Auto-contrast: ensure text is readable against fill
            var _computed_fill = shape.ComputedFill;

            if (!string.IsNullOrEmpty(_computed_fill) && ColorHelper.IsDarkColor(_computed_fill))
                text_color = "#FFFFFF";
            else if (ColorHelper.IsDarkColor(text_color))
            {
                var shape_fill = GetCellValue(shape.Cells, "FillForegnd");
                var resolved_fill = !string.IsNullOrEmpty(shape_fill) ? ColorHelper.ResolveColor(shape_fill, theme_colors) : null;

                // Also check QuickStyleFillColor for theme-based fills
                if (string.IsNullOrEmpty(resolved_fill))
                {
                    var qs_fc = GetCellValue(shape.Cells, "QuickStyleFillColor");

                    if (!string.IsNullOrEmpty(qs_fc) && theme_colors != null && theme_colors.Count > 0)
                    {
                        var qs_idx = (int)(Convert.ToSingle(qs_fc ?? "-1"));

                        if (qs_idx >= 0)
                            resolved_fill = _resolve_quickstyle_color(qs_idx, theme_colors) ?? null;
                    }
                }

                if (!string.IsNullOrEmpty(resolved_fill) && ColorHelper.IsDarkColor(resolved_fill))
                    text_color = "#FFFFFF";
            }

            var font_name = GetCellValue(char_fmt?.Cells, "Font");
            var font_family = _map_font_family(font_name);
            var style_bits = (int)(GetCellNumberValue(char_fmt?.Cells, "Style"));
            var is_bold = Convert.ToBoolean(style_bits & 1);
            var is_italic = Convert.ToBoolean(style_bits & 2);
            var is_underline = Convert.ToBoolean(style_bits & 4);

            // Paragraph alignment
            var para_fmt = shape.ParagraphFormats?.Rows?.FirstOrDefault(item => item.Index == "0");
            var halign = (int)(GetCellNumberValue(para_fmt?.Cells, "HorzAlign", 1f));
            var anchor_map = new Dictionary<int, string>() { { 0, "start" }, { 1, "middle" }, { 2, "end" } };
            var text_anchor = anchor_map.ContainsKey(halign) ? anchor_map[halign] : "middle";

            // Adjust tx for non-center horizontal alignment
            shape_left = pin_x - loc_pin_x;
            var _text_pad = font_size > 0 ? font_size * 0.4f : 4;

            if (halign == 0)
            {
                // Left-aligned: position at left edge + padding
                tx = shape_left + _text_pad;
            }
            else if (halign == 2)
            {
                // Right-aligned: position at right edge - padding
                tx = shape_left + w_px - _text_pad;
            }

            // Vertical alignment (0=top, 1=middle, 2=bottom)
            var vert_align = (int)(Convert.ToSingle(_get_cell_val(shape, "VerticalAlign", "1")));

            // Text rotation
            var txt_angle = _get_cell_float(shape, "TxtAngle");
            string txt_rotate = null;

            if (Math.Abs(txt_angle) > 1e-6)
            {
                var txt_angle_deg = -txt_angle * 180.0f / Math.PI;

                txt_rotate = $" transform=\"rotate({txt_angle_deg.ToFixed(1)},{tx.ToFixed(2)},{ty.ToFixed(2)})\"";

                if (Math.Abs(txt_angle_deg - 90) < 5 || Math.Abs(txt_angle_deg + 90) < 5)
                {
                    tx += font_size * 0.5f;
                }
            }

            // Bullet support
            var bullet = (int)(GetCellNumberValue(para_fmt?.Cells, "Bullet"));

            // Container detection for top-left label positioning
            var user_data = shape.User;
            string structure_type = GetStructureType(shape);
            bool is_container = structure_type == "Container";
            string shape_name_lower = (shape.NameU ?? shape.Name)?.ToLower();

            if (!string.IsNullOrEmpty(shape_name_lower) && (new string[] { "dash square", "container", "swimlane" }).Any(item => shape_name_lower.Contains(item)))
            {
                is_container = true;
            }

            // Font weight/style attributes
            var fw = is_bold ? " font-weight=\"bold\"" : "";
            var fs = is_italic ? " font-style=\"italic\"" : "";
            var td = is_underline ? " text-decoration=\"underline\"" : "";

            // Text width for wrapping - prefer TxtWidth over shape Width
            var txt_width = _get_cell_float(shape, "TxtWidth");
            var txt_width_px = txt_width > 0 ? txt_width * _INCH_TO_PX : w_px;

            // For very small shapes, use the larger of TxtWidth and shape Width
            if (txt_width_px < 40 && w_px > txt_width_px)
                txt_width_px = w_px;

            // For sub-shapes in groups, constrain text to actual shape width
            if (w_px > 0 && txt_width_px > w_px * 2)
                txt_width_px = w_px * 0.92f;

            // clipPath for text clipping to shape bounds (prevents overlap)
            var clip_attr = "";

            // Only clip text for shapes that are large enough to be containers
            // or have lots of text that would overflow. Small shapes should not clip.
            bool use_clip = w_px > 250 && h_px > 200;

            if (use_clip)
            {
                var clip_id = $"tclip_{shape.Id}";
                float clip_x = pin_x - loc_pin_x;
                var clip_y = pin_y - (Math.Abs(h_inch) * _INCH_TO_PX - loc_pin_y);

                // Very generous padding to avoid cutting text at edges
                float pad = font_size * 1.5f;

                lines.Add(
                $@"<defs><clipPath id=""{clip_id}""> 
                <rect x= ""{(clip_x - pad).ToFixed(2)}"" y =""{(clip_y - pad).ToFixed(2)}"" width=""{(w_px + 2 * pad).ToFixed(2)}"" height=""{(h_px + 2 * pad).ToFixed(2)}""/>
                </clipPath></defs>");

                clip_attr = $" clip-path=\"url(#{clip_id})\"";
            }

            // Container text: position at top-left
            if (is_container)
            {
                vert_align = 0;
                halign = 0;
                text_anchor = "start";
                tx = pin_x - loc_pin_x + 8;  // Left-aligned with padding
                ty = pin_y - (Math.Abs(h_inch) * _INCH_TO_PX - loc_pin_y) + font_size + 4;
            }

            // Build text lines with multi-format support
            var text_parts = shape.TextParts;

            // Check if text_parts actually have visually different formatting
            var has_multi_format = false;

            if (text_parts != null && text_parts.Count > 1 && char_formats != null)
            {
                // Compare visual properties across formats — only flag as multi-format
                // if there are real visible differences (font, color, style, size)
                var base_font = GetCellValue(char_fmt?.Cells, "Font");
                var base_color = GetCellValue(char_fmt?.Cells, "Color");
                var base_style = GetCellValue(char_fmt?.Cells, "Style", "0");
                var base_size = GetCellValue(char_fmt?.Cells, "Size");

                for (var i = 0; i < char_formats.Rows.Count; i++)
                {
                    var cp_key = char_formats.Rows[i].Index;
                    var cfmt = char_formats.Rows[i];

                    if (cp_key == "0")
                        continue;

                    var cf = GetCellValue(cfmt?.Cells, "Font");
                    var cc = GetCellValue(cfmt?.Cells, "Color");
                    var cs = GetCellValue(cfmt?.Cells, "Style", "0");
                    var csz = GetCellValue(cfmt?.Cells, "Size");

                    // Only compare fields that are set in BOTH formats
                    if (!string.IsNullOrEmpty(cf) && !string.IsNullOrEmpty(base_font) && cf != base_font)
                    {
                        has_multi_format = true;
                        break;
                    }

                    if (!string.IsNullOrEmpty(cc) && !string.IsNullOrEmpty(base_color) && cc != base_color)
                    {
                        has_multi_format = true;
                        break;
                    }

                    if (cs != base_style && cs != "Themed" && base_style != "Themed")
                    {
                        has_multi_format = true;
                        break;
                    }

                    if (!string.IsNullOrEmpty(csz) && !string.IsNullOrEmpty(base_size) && csz != base_size)
                    {
                        has_multi_format = true;
                        break;
                    }
                }
            }

            // Split text and wrap
            var text_lines = text.Split("\n");

            // Suppress master placeholder text (generic "text", "label", etc.)
            var _PLACEHOLDERS = new string[] { "text", "label", "title", "name", "value", "type" };
            var stripped = text_lines.Where(item => item.Trim().Length > 0).Select(item => item.Trim()).ToList();

            if (stripped.Count == 1 && _PLACEHOLDERS.Contains(stripped[0].ToLower()))
            {
                if (!string.IsNullOrEmpty(shape.MasterId))
                    return;
            }

            if (bullet > 0)
            {
                var bullet_char = bullet == 1 ? "• " : (bullet == 2 ? "‣ " : "– ");
                text_lines = text_lines.Select(item => !string.IsNullOrEmpty(item.Trim()) ? bullet_char + item : item).ToArray();
            }

            // Handle zero-height shapes (text-only labels in Visio)
            if (h_px <= 0)
                h_px = text_lines.Length * font_size * 1.4f;  // Estimate height from text

            // Auto-reduce font size for small shapes with long text
            if (txt_width_px > 0 && font_size > 0 && !is_container)
            {
                var avg_char_w = font_size * 0.55f;
                var total_text_len = text_lines.Sum(item => item.Length);

                // Horizontal overflow: reduce font if single-line text exceeds shape width
                var max_line_len = text_lines.Max(item => item.Length);
                var est_line_w = max_line_len * avg_char_w;

                if (est_line_w > txt_width_px * 1.1f && max_line_len > 4)
                {
                    var h_scale = txt_width_px * 0.95f / est_line_w;
                    font_size = Math.Max(5, font_size * h_scale);
                    avg_char_w = font_size * 0.55f;
                }

                // Estimate how many lines at current font size
                var est_chars_per_line = Math.Max(4, (int)(txt_width_px / avg_char_w));
                var est_lines = Math.Max(1, Math.Floor((total_text_len + est_chars_per_line - 1f) / est_chars_per_line));
                var est_text_height = est_lines * font_size * 1.2f;

                // If text would overflow height, reduce font size
                if (est_text_height > h_px * 0.9f && h_px > 0)
                {
                    var scale_factor = Math.Min(1.0f, h_px * 0.85f / est_text_height);
                    font_size = (float)Math.Max(5, font_size * scale_factor);
                }

            }

            // Word-wrap with max line limit
            if (txt_width_px > 0 && font_size > 0)
            {
                var avg_char_w = font_size * 0.55f;
                var max_chars = Math.Max(4, (int)(txt_width_px / avg_char_w));

                // For very small shapes (< 30px wide), show abbreviated text
                if (txt_width_px < 30 && !is_container)
                {
                    var first_word = text_lines != null && text_lines.Length > 0 && text_lines[0].Split(' ').Length > 0 ? text_lines[0].Split(' ')[0] : null;

                    if (!string.IsNullOrEmpty(first_word) && first_word.Length > max_chars)
                        first_word = first_word.Substring(Math.Max(2, max_chars - 1)) + "…";

                    text_lines = !string.IsNullOrEmpty(first_word) ? [first_word] : text_lines.SkipLast(1).ToArray();
                }
                else
                {
                    var wrapped_lines = new List<string>();

                    foreach (var tline in text_lines)
                    {
                        if (tline.Length <= max_chars)
                            wrapped_lines.Add(tline);
                        else
                        {
                            var words = tline.Split(' ');
                            string current = null;

                            foreach (var word in words)
                            {
                                if (!string.IsNullOrEmpty(current) && current.Length + 1 + word.Length > max_chars)
                                {
                                    wrapped_lines.Add(current);
                                    current = word;
                                }
                                else
                                    current = !string.IsNullOrEmpty(current) ? current + " " + word : word;
                            }

                            if (!string.IsNullOrEmpty(current))
                                wrapped_lines.Add(current);
                        }
                    }

                    text_lines = wrapped_lines.ToArray();
                }

                // Limit to max lines based on available height
                // First try reducing font size to fit, then truncate as last resort
                if (h_px > 0 && !is_container)
                {
                    var needed_h = text_lines.Length * font_size * 1.2f;

                    if (needed_h > h_px && font_size > 5)
                    {
                        // Try reducing font to fit all lines
                        var target_fs = h_px / (text_lines.Length * 1.2f);

                        if (target_fs >= 5)
                        {
                            font_size = target_fs;

                            // Re-wrap with new font size
                            avg_char_w = font_size * 0.55f;
                            max_chars = Math.Max(4, (int)(txt_width_px / avg_char_w));

                            var re_wrapped = new List<string>();

                            foreach (var tline in text.Split("\n"))
                            {
                                if (tline.Length <= max_chars)
                                    re_wrapped.Add(tline);
                                else
                                {
                                    var words = tline.Split(' ');
                                    string current = null;

                                    foreach (var word in words)
                                    {
                                        if (!string.IsNullOrEmpty(current) && current.Length + 1 + word.Length > max_chars)
                                        {
                                            re_wrapped.Add(current);
                                            current = word;
                                        }
                                        else
                                            current = !string.IsNullOrEmpty(current) ? current + " " + word : word;
                                    }

                                    if (!string.IsNullOrEmpty(current))
                                        re_wrapped.Add(current);
                                }
                            }

                            text_lines = re_wrapped.ToArray();
                        }
                    }
                }

                var max_lines = is_container ? 3 : (h_px > 0 ? Math.Max(2, (int)(h_px / (font_size * 1.2))) : 4);
                max_lines = Math.Min(max_lines, 8);  // absolute cap

                if (text_lines.Length > max_lines)
                {
                    text_lines = text_lines.Skip(max_lines).ToArray();

                    // Add ellipsis to last line
                    var last = text_lines.LastOrDefault();

                    if (last?.Length > 3)
                        text_lines[text_lines.Length - 1] = last.Substring(0, last.Length - 3).Trim() + "…";
                    else
                        text_lines[text_lines.Length - 1] = last + "…";
                }
            }

            var total_height = text_lines.Length * font_size * 1.2f;

            if (has_multi_format && text_parts != null && text_parts.Count > 0)
            {
                // Multi-format text: render each part as a tspan
                // Collect parts into lines by splitting on newlines
                var all_text = "";
                var part_spans = new List<dynamic>();

                foreach (var part in text_parts)
                {
                    var part_text = part.Text;

                    if (!string.IsNullOrEmpty(part_text))
                        continue;

                    var cp = part.CP ?? "0";
                    Row cfmt = char_formats?.Rows?.FirstOrDefault(item => item.Index == cp) ?? char_fmt;
                    var p_font_size = (cfmt != null ? GetCellNumberValue(cfmt.Cells, "Size") : 0.1111f) * _INCH_TO_PX;

                    if (p_font_size < 6)
                        p_font_size = 8;

                    var p_color = ColorHelper.ResolveColor(GetCellValue(cfmt.Cells, "Color"), theme_colors) ?? text_color;
                    var p_font = _map_font_family(GetCellValue(cfmt.Cells, "Font", font_name));
                    var p_style = (int)(GetCellNumberValue(cfmt.Cells, "Style"));
                    var p_bold = Convert.ToBoolean(p_style & 1) ? "bold" : "normal";
                    var p_italic = Convert.ToBoolean(p_style & 2) ? "italic" : "normal";

                    part_spans.Add(new
                    {
                        text = part_text,
                        font = p_font,
                        size = p_font_size,
                        color = p_color,
                        bold = p_bold,
                        italic = p_italic
                    });

                    all_text += part_text;
                }

                // Render as single text element with tspans
                lines.Add(
                    $@"<text x=""{tx.ToFixed(2)}"" y=""{ty.ToFixed(2)}"" text-anchor=""{text_anchor}"" dominant-baseline=""central"" font-family=""{font_family}"" font-size= ""{font_size.ToFixed(1)}"" fill=""{text_color}""{fw}{fs}{td}{txt_rotate}{clip_attr}>"
                );

                foreach (var span in part_spans)
                {
                    var escaped = _escape_xml(span["text"]);

                    // Handle newlines within spans
                    var sub_lines = escaped.split('\n');

                    for (var k = 0; k < sub_lines.Length; k++)
                    {
                        var sl = sub_lines[k];

                        if (string.IsNullOrEmpty(sl))
                            continue;

                        var dy_attr = k > 0 ? $" dy=\"{(span["size"] * 1.2f).ToFixed(1)}\" x=\"{tx.ToFixed(2)}\"" : "";

                        lines.Add(
                           $@"<tspan font-family=""{span["font"]}"" font-size=""{span["size"].ToFixed(1)}"" fill=""{span["color"]}""
                           font-weight=""{span["bold"]}"" font-style=""{span["italic"]}""{dy_attr}>
                           {sl}</tspan>"
                       );
                    }
                }

                lines.Add("</text>");
            }
            else if (text_lines.Length == 1)
            {
                // Single line
                if (!is_container)
                {
                    if (vert_align == 0)
                        ty = pin_y - h_px / 2.0f + font_size;
                    else if (vert_align == 2)
                        ty = pin_y + h_px / 2.0f - font_size * 0.3f;
                }

                // Offset connector labels above the line
                if (_is_1d_shape)
                    ty -= font_size * 0.8f + 3;

                var escaped = _escape_xml(text_lines[0]);
                // Add white background for connector labels (improves readability)
                if (_is_1d_shape && !string.IsNullOrEmpty(escaped.Trim()))
                {
                    var est_w = escaped.Length * font_size * 0.55f;
                    lines.Add(
                        $@"<rect x=""{(tx - est_w / 2 - 2).ToFixed(2)}"" y=""{(ty - font_size * 0.6).ToFixed(2)}"" width=""{(est_w + 4).ToFixed(2)}"" height=""{(font_size * 1.2).ToFixed(2)}"" fill=""white"" fill-opacity=""0.55"" rx= ""1""/>"
                    );
                }

                // Only mark as noclip if shape has visible geometry (text is inside a shape)
                // Standalone text labels should still be subject to collision avoidance
                var _has_visible_body = shape.HasGeometry && w_px > 10 && h_px > 10;
                var _noclip = (_has_visible_body && !_is_1d_shape) ? $" data-noclip=\"1\"" : "";

                lines.Add(
                    $@"<text x=""{tx.ToFixed(2)}"" y =""{ty.ToFixed(2)}"" text-anchor=""{text_anchor}"" dominant-baseline=""central"" font-family=""{font_family}"" font-size=""{font_size.ToFixed(1)}"" fill=""{text_color}""{fw}{fs}{td}{txt_rotate}{clip_attr}{_noclip}>{escaped}</text>"
                );
            }
            else
            {
                //Multi-line
                var start_y = 0.0f;
                if (!is_container)
                {
                    if (vert_align == 0)
                        start_y = pin_y - h_px / 2 + font_size;
                    else if (vert_align == 2)
                        start_y = pin_y + h_px / 2 - total_height + font_size * 0.6f;
                    else
                        start_y = ty - total_height / 2 + font_size * 0.6f;
                }
                else
                {
                    start_y = ty;
                }

                // Offset connector labels above the line
                if (_is_1d_shape)
                    start_y -= font_size * 0.8f + 3;

                for (var j = 0; j < text_lines.Length; j++)
                {
                    var tline = text_lines[j];

                    if (string.IsNullOrEmpty(tline.Trim()))
                        continue;

                    var escaped = _escape_xml(tline);
                    var ly = start_y + j * font_size * 1.2f;
                    var _has_visible_body = shape.HasGeometry && w_px > 10 && h_px > 10;
                    var _noclip = (_has_visible_body && !_is_1d_shape) ? " data-noclip=\"1\"" : "";

                    lines.Add(
                        $@"<text x=""{tx.ToFixed(2)}"" y=""{ly.ToFixed(2)}"" text-anchor=""{text_anchor}"" font-family=""{font_family}"" font-size=""{font_size.ToFixed(1)}"" fill=""{text_color}""{fw}{fs}{td}{txt_rotate}{clip_attr}{_noclip}>{escaped}</text>"
                    );
                }
            }
        }

        #endregion

        #region Page dimension parsing

        /// <summary>
        /// Convert a value in the given unit to inches.
        /// </summary>
        /// <param name="val"></param>
        /// <param name="unit"></param>
        /// <returns></returns>

        /// <summary>
        /// Normalize page dimensions to inches (in drawing coordinate space).
        /// 
        /// When PageScale and DrawingScale are present, the page dimensions are
        /// in drawing units.Shapes are also in drawing units, so we keep them
        /// consistent but ensure the pixel size is reasonable.
        /// 
        /// For scaled drawings (e.g., floorplans), the page might be 1728 FT wide
        /// but the drawing scale means shapes are positioned in those coordinates.
        /// We keep the coordinate space but cap the SVG pixel size.
        /// </summary>
        /// <param name="page_w"></param>
        /// <param name="page_h"></param>
        /// <param name="units"></param>
        /// <param name="page_scale"></param>
        /// <param name="page_scale_u"></param>
        /// <param name="draw_scale"></param>
        /// <param name="draw_scale_u"></param>
        /// <returns></returns>
        private static (float, float) _normalize_page_dims(float page_w, float page_h,
                         Dictionary<string, string> units = null,
                         float page_scale = 0.0f, string page_scale_u = null,
                         float draw_scale = 0.0f, string draw_scale_u = null)
        {
            // Visio XML stores all values in internal units (inches) regardless of
            // the U attribute (which is the display unit, not storage unit).
            // Do NOT convert by U — values are already in inches.
            return (page_w, page_h);
        }

        /// <summary>
        /// Extract page width and height from a page XML.
        /// </summary>
        /// <param name="page_xml"></param>
        /// <returns>(width_inches, height_inches).</returns>
        private static (float, float) _parse_page_dimensions(string page_xml)
        {
            var root = XDocument.Parse(page_xml).Root;

            var page_w = 8.5f;
            var page_h = 11.0f;
            var units = new Dictionary<string, string>();
            var page_scale = 0.0f;
            var draw_scale = 0.0f;
            string page_scale_u = null;
            string draw_scale_u = null;

            // Look for PageSheet
            var page_sheet = root.Child("Page").Child("PageSheet");

            if (page_sheet != null)
            {
                foreach (var cell in page_sheet.Children("Cell"))
                {
                    var n = cell.GetAttributeValue("N");
                    var v = cell.GetAttributeValue("V");
                    var u = cell.GetAttributeValue("U");

                    if (n == "PageWidth")
                    {
                        page_w = !string.IsNullOrEmpty(v) ? float.Parse(v) : 8.5f;

                        if (!string.IsNullOrEmpty(u))
                            units.Add("PageWidth", u);
                    }
                    else if (n == "PageHeight")
                    {
                        page_h = !string.IsNullOrEmpty(v) ? float.Parse(v) : 11.0f;

                        if (!string.IsNullOrEmpty(u))
                            units.Add("PageHeight", u);
                    }
                    else if (n == "PageScale")
                    {
                        page_scale = float.Parse(v);
                        page_scale_u = u;
                    }
                    else if (n == "DrawingScale")
                    {
                        draw_scale = float.Parse(v);
                        draw_scale_u = u;
                    }
                }
            }

            return _normalize_page_dims(page_w, page_h, units,
                                page_scale, page_scale_u,
                                draw_scale, draw_scale_u);
        }

        /// <summary>
        /// Parse page dimensions from pages.xml (the index file).
        /// </summary>
        /// <param name="zf"></param>
        /// <returns>List of (width_inches, height_inches) per page.
        /// Falls back to individual page XML parsing.
        /// </returns>
        private static List<(float, float)> _parse_all_page_dimensions(IArchive zf)
        {
            var dims = new List<(float, float)>();

            var pages_xml = GetFileContent(zf, "visio/pages/pages.xml");
            var root = XDocument.Parse(pages_xml).Root;

            foreach (var page in root.Children("Page"))
            {
                var pw = 8.5f;
                var ph = 11.0f;

                var units = new Dictionary<string, string>();
                var page_scale = 0.0f;
                var draw_scale = 0.0f;
                string page_scale_u = null;
                string draw_scale_u = null;
                var page_sheet = page.Child("PageSheet");

                if (page_sheet != null)
                {
                    foreach (var cell in page_sheet.Children("Cell"))
                    {
                        var n = cell.GetAttributeValue("N");
                        var v = cell.GetAttributeValue("V");
                        var u = cell.GetAttributeValue("U");

                        if (n == "PageWidth")
                        {
                            pw = !string.IsNullOrEmpty(v) ? float.Parse(v) : 8.5f;

                            if (!string.IsNullOrEmpty(u))
                                units.Add("PageWidth", u);
                        }
                        else if (n == "PageHeight")
                        {
                            ph = !string.IsNullOrEmpty(v) ? float.Parse(v) : 11.0f;

                            if (!string.IsNullOrEmpty(u))
                                units.Add("PageHeight", u);
                        }
                        else if (n == "PageScale")
                        {
                            page_scale = float.Parse(v);
                            page_scale_u = u;
                        }
                        else if (n == "DrawingScale")
                        {
                            draw_scale = float.Parse(v);
                            draw_scale_u = u;
                        }
                    }
                }

                dims.Add(_normalize_page_dims(pw, ph, units,
                                             page_scale, page_scale_u,
                                             draw_scale, draw_scale_u));
            }

            return dims;
        }
        #endregion

        #region Main parser and SVG generation

        /// <summary>
        /// Parse <Connect> elements from a page XML root.
        /// </summary>
        /// <param name="page_xml_root"></param>
        /// <returns></returns>
        private static List<Connect> _parse_connects(XElement page_xml_root)
        {
            var connects = new List<Connect>();
            var connects_el = page_xml_root.Child("Connects");

            if (connects_el == null)
            {
                return connects;
            }

            foreach (var c in connects_el.Children("Connect"))
            {
                connects.Add(new Connect()
                {
                    FromSheet = c.GetAttributeValue("FromSheet"),
                    FromCell = c.GetAttributeValue("FromCell"),
                    ToSheet = c.GetAttributeValue("ToSheet"),
                    ToCell = c.GetAttributeValue("ToCell")
                });
            }

            return connects;
        }

        /// <summary>
        /// Build a flat index of shape ID -> shape dict, including sub-shapes.
        /// </summary>
        /// <param name="shapes"></param>
        /// <returns></returns>
        private static Dictionary<string, Shape> _build_shape_index(List<Shape> shapes)
        {
            var idx = new Dictionary<string, Shape>();

            foreach (var s in shapes)
            {
                idx.Add(s.Id, s);

                foreach (var sub in s.SubShapes)
                {
                    idx[sub.Id] = sub;

                    // Also index deeper sub-shapes
                    foreach (var subsub in sub.SubShapes)
                    {
                        idx[subsub.Id] = subsub;
                    }
                }
            }

            return idx;
        }

        /// <summary>
        /// Resolve a connection cell reference to page coordinates (px).
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="cell_ref">like 'Controls.Row_1' or 'Connections.X1'</param>
        /// <param name="page_h"></param>
        /// <param name="shape_index"></param>
        /// <returns></returns>
        private static (float?, float?) _resolve_connection_point(Shape shape, string cell_ref, float page_h, Dictionary<string, Shape> shape_index)
        {
            var pin_x = _get_cell_float(shape, "PinX");
            var pin_y = _get_cell_float(shape, "PinY");
            var loc_pin_x = _get_cell_float(shape, "LocPinX");
            var loc_pin_y = _get_cell_float(shape, "LocPinY");

            if (cell_ref.StartsWith("Controls."))
            {
                var row_key = cell_ref.Split('.')[1];  // e.g. "Row_1"
                Section ctrl = shape.Controls.FirstOrDefault(item => item.Index == row_key);

                if (ctrl != null)
                {
                    var lx = GetSectionRowCellNumberValue(ctrl, "X");
                    var ly = GetSectionRowCellNumberValue(ctrl, "Y");

                    // Local to page
                    var px = (pin_x - loc_pin_x + lx) * _INCH_TO_PX;
                    var py = (page_h - (pin_y - loc_pin_y + ly)) * _INCH_TO_PX;

                    return (px, py);
                }
            }
            else if (cell_ref.StartsWith("Connections."))
            {
                // Parse "X1" -> row IX=0, "X2" -> IX=1, etc
                var suffix = cell_ref.Split('.')[1]; // e.g. "X1"
                var m = Regex.Match(suffix, @"X(\d+)");

                if (m.Success)
                {
                    var row_ix = (int.Parse(m.Groups[1].Value) - 1).ToString(); //  # X1 -> IX=0
                    Section conn = shape.Connections.FirstOrDefault(item => item.Index == row_ix);

                    if (conn != null)
                    {
                        var lx = GetSectionRowCellNumberValue(conn, "X");
                        var ly = GetSectionRowCellNumberValue(conn, "Y");
                        var px = (pin_x - loc_pin_x + lx) * _INCH_TO_PX;
                        var py = (page_h - (pin_y - loc_pin_y + ly)) * _INCH_TO_PX;

                        return (px, py);
                    }
                }
            }

            return default((float?, float?));
        }

        /// <summary>
        /// Clip a line endpoint (x2,y2) to the edge of a rectangle centered at (cx,cy).
        /// 
        /// hw, hh = half-width, half-height.
        /// </summary>
        /// <param name="x1"></param>
        /// <param name="y1"></param>
        /// <param name="x2"></param>
        /// <param name="y2"></param>
        /// <param name="cx"></param>
        /// <param name="cy"></param>
        /// <param name="hw"></param>
        /// <param name="hh"></param>
        /// <returns>The clipped (x2, y2) on the rectangle edge.</returns>
        private static (float, float) _clip_line_to_rect(float x1, float y1, float x2, float y2, float cx, float cy, float hw, float hh)
        {
            var dx = x2 - cx;
            var dy = y2 - cy;

            if (Math.Abs(dx) < 0.01f && Math.Abs(dy) < 0.01f)
                return (x2, y2);

            // Direction from center to the other point
            var ddx = x1 - cx;
            var ddy = y1 - cy;

            if (Math.Abs(ddx) < 0.01f && Math.Abs(ddy) < 0.01f)
                return (x2, y2);

            // Find intersection with rectangle edges
            var t_vals = new List<float>();

            if (Math.Abs(ddx) > 0.01f)
            {
                t_vals.Add(hw / Math.Abs(ddx));
                t_vals.Add(-hw / Math.Abs(ddx));
            }

            if (Math.Abs(ddy) > 0.01f)
            {
                t_vals.Add(hh / Math.Abs(ddy));
                t_vals.Add(-hh / Math.Abs(ddy));
            }

            var t = float.MaxValue;

            foreach (var tv in t_vals)
            {
                if (tv > 0)
                {
                    var px = cx + ddx * tv;
                    var py = cy + ddy * tv;

                    if (Math.Abs(px - cx) <= hw + 0.5f && Math.Abs(py - cy) <= hh + 0.5f)
                        t = Math.Min(t, tv);
                }
            }

            if (t < float.MaxValue)
                return (cx + ddx * t, cy + ddy * t);

            return (x2, y2);
        }

        /// <summary>
        /// Render connection lines as SVG elements.
        /// </summary>
        /// <param name="masters"></param>
        /// <returns></returns>
        private static List<string> _render_connections_svg(List<Connect> connects, Dictionary<string, Shape> shape_index,
                            float page_h, Dictionary<string, Dictionary<string, Shape>> masters)
        {
            var lines = new List<string>();

            // Track which connector shapes (1D) already render their own lines
            var connector_sheets = new HashSet<string>();

            foreach (var conn in connects)
            {
                var from_cell = conn.FromCell;

                if (from_cell == "BeginX" || from_cell == "EndX")
                    connector_sheets.Add(conn.FromSheet);
            }

            foreach (var conn in connects)
            {
                var from_shape = shape_index.ContainsKey(conn.FromSheet) ? shape_index[conn.FromSheet] : null;
                var to_shape = shape_index.ContainsKey(conn.ToSheet) ? shape_index[conn.ToSheet] : null;

                if (from_shape == null || to_shape == null)
                    continue;

                // Skip if the from_sheet is a 1D connector (already rendered)
                if (connector_sheets.Contains(conn.FromSheet))
                    continue;

                // Merge with masters for connections/controls data
                from_shape = _merge_shape_with_master(from_shape, masters, from_shape.MasterId);
                to_shape = _merge_shape_with_master(to_shape, masters, to_shape.MasterId);

                (float?, float?) from_pt = _resolve_connection_point(from_shape, conn.FromCell, page_h, shape_index);
                (float?, float?) to_pt = _resolve_connection_point(to_shape, conn.ToCell, page_h, shape_index);

                if (from_pt.Item1.HasValue && to_pt.Item1.HasValue)
                {
                    // Clip endpoints to shape edges
                    var fw = Math.Abs(_get_cell_float(from_shape, "Width")) * _INCH_TO_PX / 2.0f;
                    var fh = Math.Abs(_get_cell_float(from_shape, "Height")) * _INCH_TO_PX / 2.0f;
                    var tw = Math.Abs(_get_cell_float(to_shape, "Width")) * _INCH_TO_PX / 2.0f;
                    var th = Math.Abs(_get_cell_float(to_shape, "Height")) * _INCH_TO_PX / 2.0f;

                    if (fw > 5 && fh > 5)
                    {
                        var fcx = (_get_cell_float(from_shape, "PinX")) * _INCH_TO_PX;
                        var fcy = (page_h - _get_cell_float(from_shape, "PinY")) * _INCH_TO_PX;
                        from_pt = _clip_line_to_rect(to_pt.Item1.Value, to_pt.Item2.Value, from_pt.Item1.Value, from_pt.Item2.Value, fcx, fcy, fw, fh);
                    }

                    if (tw > 5 && th > 5)
                    {
                        var tcx = (_get_cell_float(to_shape, "PinX")) * _INCH_TO_PX;
                        var tcy = (page_h - _get_cell_float(to_shape, "PinY")) * _INCH_TO_PX;
                        to_pt = _clip_line_to_rect(from_pt.Item1.Value, from_pt.Item2.Value, to_pt.Item1.Value, to_pt.Item2.Value, tcx, tcy, tw, th);
                    }

                    lines.Add(
                       $@"<line x1=""{from_pt.Item1.Value.ToFixed(2)}"" y1=""{from_pt.Item2.Value.ToFixed(2)}"" x2=""{to_pt.Item1.Value.ToFixed(2)}"" y2=""{to_pt.Item2.Value.ToFixed(2)}""
                       stroke=""#555555"" stroke-width=""1.50"" marker-end=""url(#arrow_end_2_555555)""/>"
                   );
                }
                else if (to_pt.Item1.HasValue)
                {
                    float bus_y = (page_h - _get_cell_float(from_shape, "PinY")) * _INCH_TO_PX;

                    lines.Add(
                        $@"<line x1=""{to_pt.Item1.Value.ToFixed(2)}"" y1=""{bus_y.ToFixed(2)}""  x2=""{to_pt.Item1.Value.ToFixed(2)}"" y2=""{to_pt.Item2.Value.ToFixed(2)}"" stroke=""#555555"" stroke-width=""1.50""/>"
                    );
                }
            }

            return lines;
        }

        /// <summary>
        /// Parse shapes from a Visio page XML into rich shape dicts.
        /// </summary>
        /// <param name="page_xml">Raw XML bytes of a page file.</param>
        /// <param name="master_texts">Legacy param (ignored, kept for API compat).</param>
        /// <param name="masters">Full master shapes dict from _parse_master_shapes.</param>
        /// <returns></returns>
        private static List<Shape> _parse_vsdx_shapes(string page_xml, Dictionary<string, string> master_texts = null, Dictionary<string, Dictionary<string, Shape>> masters = null)
        {
            var shapes = new List<Shape>();

            var root = XDocument.Parse(page_xml).Root;

            // Find all top-level shapes (direct children of Shapes element)
            var shapes_container = root.Child("Shapes");

            if (shapes_container == null)
                return shapes;

            foreach (var shape_elem in shapes_container.Children("Shape"))
            {
                var sd = _parse_single_shape(shape_elem);

                shapes.Add(sd);
            }

            return shapes;
        }

        /// <summary>
        /// Shift overlapping labels and optionally add background for readability.
        /// 
        /// Only adds white background rectangles when text actually collides with
        /// other text, avoiding visual noise in sparse diagrams.
        /// </summary>
        /// <param name="text_elements"></param>
        /// <returns></returns>
        private static List<string> _avoid_text_collisions(List<string> text_elements)
        {
            // First pass: parse all text elements to build boxes
            var parsed = new List<(string, Dictionary<string, dynamic>)>(); //list of (elem, tx, ty, fs, clean_txt, est_w, est_h, box_x, box_y, is_noclip)

            Regex _text_re = new Regex(@"<text\s+x=""([^""]+)""\s+y=""([^""]+)""[^>]*font-size=""([^""]+)""[^>]*>(.*?)</text>");

            foreach (var elem in text_elements)
            {
                var m = _text_re.Match(elem);

                if (!m.Success || elem.Contains("data-noclip=\"1\""))
                {
                    parsed.Add((elem, null));
                    continue;
                }

                var tx = float.Parse(m.Groups[1].Value);
                var ty = float.Parse(m.Groups[2].Value);
                var fs = float.Parse(m.Groups[3].Value);
                var txt = m.Groups[4].Value;

                var clean_txt = Regex.Replace(txt, @"<[^>]+>", "");

                if (string.IsNullOrEmpty(clean_txt.Trim()))
                {
                    parsed.Add((elem, null));
                    continue;
                }

                var est_w = clean_txt.Length * fs * 0.55f;
                var est_h = fs * 1.3f;

                var box_x = 0.0f;

                if (elem.Contains("text-anchor=\"start\""))
                    box_x = tx - 1;
                else if (elem.Contains("text-anchor=\"end\""))
                    box_x = tx - est_w - 1;
                else
                    box_x = tx - est_w / 2.0f - 1;

                var box_y = ty - est_h * 0.55f;

                parsed.Add((elem, new Dictionary<string, dynamic>()
                {
                    // type: ignore[arg-type]
                    {"tx" , tx },
                    {"ty", ty },
                    {"fs" , fs },
                    {"orig_y", m.Groups[2].Value },
                    {"clean_txt", clean_txt },
                    {"est_w", est_w },
                    {"est_h", est_h },
                    {"box_x" , box_x },
                    {"box_y" , box_y }
                }));
            }

            // Determine diagram density for collision strategy tuning
            var text_count = parsed.Where(item => item.Item2 != null).Count();
            bool is_dense = text_count > 40; // dense diagram threshold

            // For very dense diagrams, scale down font sizes to reduce collisions
            if (text_count > 60)
            {
                var _scale = Math.Max(0.65, 40.0 / text_count);

                for (var i = 0; i < parsed.Count; i++)
                {
                    var item = parsed[i];
                    var elem = item.Item1;
                    var data = item.Item2;

                    if (data != null)
                    {
                        data["fs"] *= _scale;
                        data["est_w"] = (data["clean_txt"]).Length * data["fs"] * 0.55f;
                        data["est_h"] = data["fs"] * 1.3f;

                        if (elem.Contains("text-anchor=\"start\""))
                            data["box_x"] = data["tx"] - 1;
                        else if (elem.Contains("text-anchor=\"end\""))
                            data["box_x"] = data["tx"] - data["est_w"] - 1;
                        else
                        {
                            data["box_x"] = data["tx"] - data["est_w"] / 2.0f - 1;
                            data["box_y"] = data["ty"] - data["est_h"] * 0.55f;
                        }

                        // Update SVG element with scaled font size
                        var _orig_fs_str = $"font-size=\"{((float)(data["fs"] / _scale)).ToFixed(2)}\"";
                        var _new_fs_str = $"font-size=\"{data["fs"].ToFixed(2)}\"";

                        parsed[i] = (elem.Replace(_orig_fs_str, _new_fs_str), data);
                    }
                }
            }

            var placed_boxes = new List<(float, float, float, float)>();
            var collided_indices = new HashSet<int>();
            var result = new List<string>();

            for (var idx = 0; idx < parsed.Count; idx++)
            {
                var item = parsed[idx];
                var elem = item.Item1;
                var data = item.Item2;

                if (data == null)
                {
                    result.Add(elem);
                    continue;
                }

                var tx = data["tx"];
                var ty = data["ty"];
                var fs = data["fs"];
                var est_w = data["est_w"];
                var est_h = data["est_h"];
                var box_x = data["box_x"];
                var box_y = data["box_y"];

                // Collision detection: shift text to avoid overlaps
                // For dense diagrams, try shifting in multiple directions
                var max_attempts = is_dense ? 5 : 3;
                var had_collision = false;
                dynamic best_shift = null;

                for (var attempt = 0; attempt < max_attempts; attempt++)
                {
                    var collision = false;

                    for (var j = 0; j < placed_boxes.Count; j++)
                    {
                        var (px, py, pw, ph) = placed_boxes[j];

                        var overlap_x = Math.Min(box_x + est_w + 2, px + pw) - Math.Max(box_x, px);
                        var overlap_y = Math.Min(box_y + est_h, py + ph) - Math.Max(box_y, py);

                        if (overlap_x > 0 && overlap_y > 0)
                        {
                            // Only shift if overlap is significant (>20% of text height)
                            if (overlap_y > est_h * 0.2f)
                            {
                                collision = true;
                                had_collision = true;
                                collided_indices.Add(j);

                                // Alternate shift direction: down first, then try right
                                if (attempt % 2 == 0)
                                {
                                    ty += est_h + 2;
                                    box_y = ty - est_h * 0.55f;
                                }
                                else
                                {
                                    tx += est_w * 0.3f;

                                    if (elem.Contains("text-anchor=\"start\""))
                                        box_x = tx - 1;
                                    else if (elem.Contains("text-anchor=\"end\""))
                                        box_x = tx - est_w - 1;
                                    else
                                        box_x = tx - est_w / 2.0f - 1;
                                }

                                break;
                            }
                        }
                    }

                    if (!collision)
                    {
                        break;
                    }
                }

                if (had_collision)
                    collided_indices.Add(placed_boxes.Count);

                placed_boxes.Add((box_x, box_y, est_w + 2, est_h));

                // Update y if shifted
                var updated_elem = elem;
                var orig_y = data["orig_y"];

                if (Math.Abs(ty - float.Parse(orig_y)) > 0.5f)
                {
                    updated_elem = updated_elem.Replace($"y=\"{orig_y}\"", $"y=\"{ty.ToFixed(2)}\"");
                }

                // Only add white background when this text actually collided with something
                // This avoids visual noise from unnecessary white rects
                if (had_collision && fs >= 7.5f)
                {
                    string text_fill = null;
                    var _fill_m = Regex.Match(elem, @"fill=""([^""]+)""");

                    if (_fill_m.Success)
                        text_fill = _fill_m.Groups[1].Value.Trim().ToUpper();

                    var _is_light_text = false;

                    if (text_fill.StartsWith("#") && text_fill.Length == 7)
                    {
                        var _r = Convert.ToInt16(text_fill.Substring(1, 2));
                        var _g = Convert.ToInt16(text_fill.Substring(3, 2));
                        var _b = Convert.ToInt16(text_fill.Substring(5, 2));

                        var _lum = _r * 0.299f + _g * 0.587f + _b * 0.114f;
                        _is_light_text = _lum > 160;
                    }
                    else if (text_fill == "WHITE" || text_fill == "#FFF" || text_fill == "#FFFFFF")
                    {
                        _is_light_text = true;
                    }

                    if (!_is_light_text)
                    {
                        result.Add(
                           $@"<rect x=""{box_x.ToFixed(2)}"" y=""{box_y.ToFixed(2)}"" width=""{(est_w + 4).ToFixed(2)}"" height=""{est_h.ToFixed(2)}"" fill=""white"" fill-opacity=""0.7"" rx=""2""/>"
                       );
                    }
                }

                result.Add(updated_elem);
            }

            return result;
        }

        /// <summary>
        /// Generate SVG string from parsed shapes.
        /// </summary>
        /// <param name="shapes"></param>
        /// <param name="page_w"></param>
        /// <param name="page_h"></param>
        /// <param name="masters"></param>
        /// <param name="connects"></param>
        /// <param name="media"></param>
        /// <param name="page_rels"></param>
        /// <param name="bg_shapes"></param>
        /// <param name="bg_connects"></param>      
        /// <param name="theme_colors"></param>
        /// <param name="layers"></param>
        /// <returns></returns>
        private static string _shapes_to_svg(List<Shape> shapes, float page_w, float page_h, Document document,
                   Dictionary<string, Dictionary<string, Shape>> masters = null,
                   List<Connect> connects = null,
                   Dictionary<string, byte[]> media = null,
                   Dictionary<string, string> page_rels = null,
                   List<Shape> bg_shapes = null,
                   List<Connect> bg_connects = null,
                   Dictionary<string, string> theme_colors = null,
                   Dictionary<string, Layer> layers = null)
        {
            var page_w_px = page_w * _INCH_TO_PX;
            var page_h_px = page_h * _INCH_TO_PX;

            // Compute content bounding box for optimal viewBox.
            // This prevents clipped content (shapes near page edges) and removes
            // excessive whitespace (content only using part of the page).
            var vb_x = 0.0f;
            var vb_y = 0.0f;
            var vb_w = page_w_px;
            var vb_h = page_h_px;
            var max_svg_px = 4000.0f;

            var all_shapes = shapes.Select(item => item).ToList();

            if (bg_shapes != null && bg_shapes.Count > 0)
                all_shapes.AddRange(bg_shapes);

            if (all_shapes != null && all_shapes.Count > 0)
            {
                var min_x = float.MaxValue;
                var min_y = float.MaxValue;
                var max_x = float.MinValue;
                var max_y = float.MinValue;

                foreach (var s in all_shapes)
                {
                    var px = GetCellNumberValue(s.Cells, "PinX") * _INCH_TO_PX;
                    var py = (page_h - GetCellNumberValue(s.Cells, "PinY")) * _INCH_TO_PX;
                    var sw = Math.Abs(GetCellNumberValue(s.Cells, "Width")) * _INCH_TO_PX;
                    var sh = Math.Abs(GetCellNumberValue(s.Cells, "Height")) * _INCH_TO_PX;

                    if (px > 0 || py > 0)
                    {
                        min_x = Math.Min(min_x, px - sw / 2.0f);
                        min_y = Math.Min(min_y, py - sh / 2.0f);
                        max_x = Math.Max(max_x, px + sw / 2.0f);
                        max_y = Math.Max(max_y, py + sh / 2.0f);
                    }

                    // Also account for text below shapes (TxtPinY < 0)
                    var txt_pin_y = GetCellNumberValue(s.Cells, "TxtPinY");

                    if (txt_pin_y < 0)
                    {
                        // Text extends below shape
                        var text_below = Math.Abs(txt_pin_y) * _INCH_TO_PX + 20;  // font estimate
                        max_y = Math.Max(max_y, py + sh / 2.0f + text_below);
                    }

                    // Account for text that may overflow shape bounds
                    // (e.g. bullet lists, long descriptions next to shapes)
                    var txt_w = Math.Abs(GetCellNumberValue(s.Cells, "TxtWidth")) * _INCH_TO_PX;
                    var txt_h = Math.Abs(GetCellNumberValue(s.Cells, "TxtHeight")) * _INCH_TO_PX;

                    if (txt_w > sw)
                    {
                        max_x = Math.Max(max_x, px + txt_w / 2.0f);
                        min_x = Math.Min(min_x, px - txt_w / 2.0f);
                    }

                    if (txt_h > sh)
                        max_y = Math.Max(max_y, py + txt_h / 2.0f);
                }

                if (min_x < float.MaxValue)
                {
                    // Add padding — 4% or at least 20px for margins
                    var content_w = max_x - min_x;
                    var content_h = max_y - min_y;
                    var pad_x = Math.Max(50, content_w * 0.08f);
                    var pad_y = Math.Max(50, content_h * 0.08f);

                    // Use content bounds but don't shrink below page if content fills it
                    vb_x = Math.Min(0, min_x - pad_x);
                    vb_y = Math.Min(0, min_y - pad_y);
                    vb_w = content_w > page_w_px * 0.8f ? Math.Max(content_w + 2 * pad_x, page_w_px) : Math.Max(content_w + 2 * pad_x, page_w_px * 0.5f);
                    vb_h = content_h > page_h_px * 0.8f ? Math.Max(content_h + 2 * pad_y, page_h_px) : Math.Max(content_h + 2 * pad_y, page_h_px * 0.5f);

                    // Ensure viewBox covers from vb_x to max content + padding
                    vb_w = Math.Max(vb_w, max_x + pad_x - vb_x);
                    vb_h = Math.Max(vb_h, max_y + pad_y - vb_y);
                }
            }

            // Cap display pixel size
            var display_w = vb_w;
            var display_h = vb_h;

            if (Math.Max(vb_w, vb_h) > max_svg_px)
            {
                var scale = max_svg_px / Math.Max(vb_w, vb_h);
                display_w = vb_w * scale;
                display_h = vb_h * scale;
            }

            var svg_lines = new List<string>()
            {
                @"<?xml version=""1.0"" encoding=""UTF-8""?>",
                $@"<svg xmlns=""http://www.w3.org/2000/svg"" xmlns:xlink=""http://www.w3.org/1999/xlink"" width=""{display_w.ToFixed(0)}"" height=""{display_h.ToFixed(0)}"" viewBox=""{vb_x.ToFixed(0)} {vb_y.ToFixed(0)} {vb_w.ToFixed(0)} {vb_h.ToFixed(0)}"">",
                $@"<rect x=""{vb_x.ToFixed(0)}"" y=""{vb_y.ToFixed(0)}"" width=""{vb_w.ToFixed(0)}"" height=""{vb_h.ToFixed(0)}"" fill=""white""/>"
            };

            if (masters == null)
                masters = new Dictionary<string, Dictionary<string, Shape>>();
            if (media == null)
                media = new Dictionary<string, byte[]>();
            if (page_rels == null)
                page_rels = new Dictionary<string, string>();
            if (theme_colors == null)
                theme_colors = new Dictionary<string, string>();
            if (layers == null)
                layers = new Dictionary<string, Layer>();

            var used_markers = new HashSet<string>();
            var gradients = new Dictionary<string, Gradient>();
            var has_shadow = new HashSet<string>();

            // Two-pass rendering: geometry first, then text on top
            var text_layer = new List<string>();

            // Render background page shapes first (behind foreground)
            if (bg_shapes != null && bg_shapes.Count > 0)
            {
                svg_lines.Add("<!-- Background page -->");

                foreach (var s in bg_shapes)
                {
                    var svg_elements = _render_shape_svg(
                          s, page_h, document, masters, null, 0, media,
                          page_rels, used_markers = used_markers,
                          theme_colors = theme_colors,
                          layers = layers, gradients = gradients, has_shadow = has_shadow,
                          text_layer = text_layer);


                    svg_lines.AddRange(svg_elements);
                }

                if (bg_connects != null && bg_connects.Count > 0)
                {
                    var bg_index = _build_shape_index(bg_shapes);

                    svg_lines.AddRange(_render_connections_svg(
                        bg_connects, bg_index, page_h, masters));
                }
            }

            // Sort shapes: containers first (background), then regular shapes
            // This ensures containers don't obscure the shapes inside them
            //Return sort key: containers get low values (render first/behind).
            Func<Shape, int> _shape_z_order = (s) =>
            {
                var user = s.User;
                var st = GetStructureType(s);
                string name = (s.NameU ?? s.Name)?.ToLower();

                if (st == "Container" || (!string.IsNullOrEmpty(name) && (new string[] { "dash square", "container", "swimlane" }).Any(item => name.Contains(item))))
                    return 0;  // Containers render first (behind everything)

                return 1;  // Regular shapes render on top
            };

            var sorted_shapes = shapes.OrderBy(item => _shape_z_order(item));

            // Render foreground shapes (geometry only, text collected separately)
            foreach (var s in sorted_shapes)
            {
                var svg_elements = _render_shape_svg(
                                s, page_h, document, masters, null, 0, media = media,
                                page_rels = page_rels, used_markers = used_markers,
                                theme_colors = theme_colors,
                                layers = layers, gradients = gradients, has_shadow = has_shadow,
                                text_layer = text_layer);

                svg_lines.AddRange(svg_elements);
            }

            // Render connections
            if (connects != null && connects.Count > 0)
            {
                var shape_index = _build_shape_index(shapes);
                var conn_lines = _render_connections_svg(connects, shape_index, page_h, masters);

                svg_lines.AddRange(conn_lines);
            }

            // Text layer on top of everything — with collision avoidance
            if (text_layer != null && text_layer.Count > 0)
            {
                svg_lines.Add("<!-- Text layer -->");

                text_layer = _avoid_text_collisions(text_layer);
                svg_lines.AddRange(text_layer);
            }

            svg_lines.Add("</svg>");

            // Build a single <defs> block with all definitions
            var defs_content = new List<string>();

            if (used_markers != null && used_markers.Count > 0)
            {
                // _arrow_marker_defs returns ["<defs>", ...markers..., "</defs>"]
                var marker_lines = _arrow_marker_defs(used_markers);

                // Extract content between <defs> and </defs>
                foreach (var ml in marker_lines)
                    if (ml.Trim() != "<defs>" && ml.Trim() != "</defs>")
                        defs_content.Add(ml);
            }

            if (gradients != null && gradients.Count > 0)
                defs_content.AddRange(_gradient_defs(gradients));
            if (has_shadow != null && has_shadow.Count > 0)
                defs_content.Add(_shadow_filter_def());

            if (defs_content != null && defs_content.Count > 0)
            {
                var defs_lines = Enumerable.Concat(Enumerable.Concat(["<defs>"], defs_content), ["</defs>"]).ToArray();

                for (var j = 0; j < defs_lines.Length; j++)
                {
                    var ml = defs_lines[j];
                    svg_lines.Insert(3 + j, ml);
                }
            }

            return string.Join("\n", svg_lines);
        }

        private static string GetCellValue(List<Cell> cells, string name, string defaultValue = null)
        {
            if (cells == null)
            {
                return null;
            }

            var cell = cells.FirstOrDefault(item => item.Name == name);

            if (cell != null)
            {
                return cell.Value ?? defaultValue;
            }

            return defaultValue;
        }

        private static bool HasCell(List<Cell> cells, string name)
        {
            var cell = cells.FirstOrDefault(item => item.Name == name);

            return cell != null;
        }

        private static float GetCellNumberValue(List<Cell> cells, string name, float defaultValue = 0.0f)
        {
            var value = GetCellValue(cells, name);

            if (!string.IsNullOrEmpty(value))
            {
                if (float.TryParse(value, out var val))
                {
                    return val;
                }
            }

            return defaultValue;
        }

        private static string GetCellFormular(List<Cell> cells, string name)
        {
            var cell = cells.FirstOrDefault(item => item.Name == name);

            if (cell != null)
            {
                return cell.Formula;
            }

            return null;
        }

        private static float GetSectionRowCellNumberValue(Section section, string cellName, float defaultValue = 0.0f)
        {
            if (section == null || section.Rows == null)
            {
                return defaultValue;
            }

            foreach (var row in section.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.Name == cellName)
                    {
                        if (float.TryParse(cell.Value, out var val))
                        {
                            return val;
                        }
                    }
                }
            }

            return defaultValue;
        }

        private static string GetStructureType(Shape shape)
        {
            var userSection = shape.User;

            return userSection?.Rows?.FirstOrDefault(item => item.Name == "msvStructureType")?.Cells?.FirstOrDefault(item => item.Name == "Value")?.Value;
        }
        #endregion

        #region Public API functions      

        /// <summary>
        /// Get sorted list of page XML files from a ZIP.
        /// </summary>
        /// <param name="zf"></param>
        /// <returns></returns>
        private static List<string> _get_page_files(IArchive zf)
        {
            string pagesPath = "visio/pages/pages.xml";
            string relsPath = "visio/pages/_rels/pages.xml.rels";

            string pagesContent = GetFileContent(zf, pagesPath);
            string relsContent = GetFileContent(zf, relsPath);

            XElement pagesRoot = XDocument.Parse(pagesContent).Root;
            XElement relsRoot = XDocument.Parse(relsContent).Root;

            Dictionary<string, string> rels = new Dictionary<string, string>();

            foreach (var rel in relsRoot.Children("Relationship"))
            {
                var rid = rel.GetAttributeValue("Id");
                var target = rel.GetAttributeValue("Target");

                rels.Add(rid, target);
            }

            List<string> pagePaths = new List<string>();

            foreach (var page in pagesRoot.Children("Page"))
            {
                bool isCustomName = page.GetAttributeValue("IsCustomName") == "1";

                if (!isCustomName)
                {
                    string rid = page.Child("Rel")?.GetAttributeValue("id");

                    if (!string.IsNullOrEmpty(rid) && rels.ContainsKey(rid))
                    {
                        pagePaths.Add($"visio/pages/{rels[rid]}");
                    }
                }
            }

            if (pagePaths.Count == 0)
            {
                pagePaths = zf.Entries.Where(item => item.Key.ToLower().Contains("page") && item.Key.EndsWith(".xml") && !item.Key.ToLower().Contains("pages.xml") && !item.Key.Contains("_rels"))
                .OrderBy(item => item.Key, StringComparison.OrdinalIgnoreCase.WithNaturalSort())
                .Select(item => item.Key).ToList();
            }

            return pagePaths;
        }

        /// <summary>
        /// Parse pages.xml to find background page references.
        /// </summary>
        /// <param name="zf"></param>
        /// <returns>{page_index: background_page_index} (0-based).</returns>
        private static Dictionary<int, int> _parse_background_pages(IArchive zf)
        {
            var bg_map = new Dictionary<int, int>();

            var pages_xml = GetFileContent(zf, "visio/pages/pages.xml");
            var root = XDocument.Parse(pages_xml).Root;
            var pages = root.Children("Page");

            // Build page ID -> index map
            var page_id_to_idx = new Dictionary<string, int>();

            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                var pid = page.GetAttributeValue("ID");

                if (!string.IsNullOrEmpty(pid))
                    page_id_to_idx.Add(pid, i);
            }

            // Find BackPage references
            for (int i = 0; i < pages.Count; i++)
            {
                var page = pages[i];
                var page_sheet = page.Child("PageSheet");

                if (page_sheet == null)
                {
                    continue;
                }

                foreach (var cell in page_sheet.Children("Cell"))
                {
                    if (cell.GetAttributeValue("N") == "BackPage")
                    {
                        var back_id = cell.GetAttributeValue("V");

                        if (!string.IsNullOrEmpty(back_id) && page_id_to_idx.ContainsKey(back_id))
                            bg_map.Add(i, page_id_to_idx[back_id]);
                    }
                }
            }

            return bg_map;
        }

        /// <summary>
        /// Parse .vsdx/.vstx/.vssx (ZIP+XML) and generate SVG directly.
        /// </summary>
        /// <param name="input_path"></param>
        /// <param name="output_dir"></param>
        /// <returns></returns>
        private static List<string> _vsdx_to_svg(string input_path, string output_dir)
        {
            if (!Directory.Exists(output_dir))
            {
                Directory.CreateDirectory(output_dir);
            }

            string basename = Path.GetFileNameWithoutExtension(input_path);

            List<string> svgs = ConvertToSvg(input_path);

            List<string> svg_files = new List<string>();

            for (int i = 0; i < svgs.Count; i++)
            {
                string svg_content = svgs[i];

                var svg_path = Path.Combine(output_dir, $"{basename}_page{i + 1}.svg");

                System.IO.File.WriteAllText(svg_path, svg_content, Encoding.UTF8);

                svg_files.Add(svg_path);
            }

            return svg_files;
        }

        /// <summary>
        /// Convert a Visio file to SVG pages.
        /// </summary>
        /// <param name="input_path"></param>
        /// <param name="output_dir"></param>
        /// <returns>A list of SVG file paths (one per page).</returns>
        public static List<string> convert_vsd_to_svg(string input_path, string output_dir)
        {
            if (!Directory.Exists(output_dir))
            {
                Directory.CreateDirectory(output_dir);
            }

            string ext = Path.GetExtension(input_path).ToLower();

            if (!ALL_EXTENSIONS.Contains(ext))
            {
                throw new NotSupportedException($"Unsupported file format:{ext}");
            }

            // For .vsdx files, prefer built-in parser (handles images, arrows,
            // background pages). Fall back to libvisio only for .vsd/.vss.
            if (_XML_EXTENSIONS.Contains(ext))
            {
                var svg_files = _vsdx_to_svg(input_path, output_dir);

                if (svg_files != null && svg_files.Count > 0)
                    return svg_files;
            }

            // For .vsd binary files, try native parser first
            if (Array.Exists(new string[] { ".vsd", ".vss", ".vst" }, item => item == ext))
            {
                //var svg_files = _vsd_to_svg(input_path, output_dir);

                //if (svg_files != null && svg_files.Count > 0)
                //    return svg_files;

                throw new NotImplementedException("Not implemented yet.");
            }

            return null;
        }

        public static List<string> ConvertToSvg(string filePath)
        {
            var svgs = new List<string>();

            IArchive zf = null;

            try
            {
                zf = ArchiveFactory.OpenArchive(filePath, GetZipOptions());
            }
            catch (Exception ex)
            {
                return svgs;
            }

            using (zf)
            {
                var masters = _parse_master_shapes(zf);
                var media = _extract_media(zf);
                var theme_colors = _parse_theme(zf);
                var doc = _parse_document(zf);

                // Parse master rels for ForeignData image resolution
                var master_rels = new Dictionary<string, string>();

                foreach (var name in zf.Entries.Select(item => item.Key))
                {
                    if (name.StartsWith("visio/masters/_rels/master") && name.EndsWith(".xml.rels"))
                    {
                        var rels_xml = GetFileContent(zf, name);
                        var root = XDocument.Parse(rels_xml).Root;

                        foreach (var rel in root.Children("Relationship"))
                        {
                            var rid = rel.GetAttributeValue("Id");
                            var target = rel.GetAttributeValue("Target");

                            if (!string.IsNullOrEmpty(rid) && !string.IsNullOrEmpty(target))
                                master_rels[rid] = target;
                        }
                    }
                }

                var page_files = _get_page_files(zf);
                var all_dims = _parse_all_page_dimensions(zf);
                var bg_map = _parse_background_pages(zf);

                // Pre-parse all pages for background composition
                // idx -> (shapes, connects, page_rels, layers)
                var page_cache = new Dictionary<int, (List<Shape> shapes, List<Connect> connects, Dictionary<string, string> page_rels, Dictionary<string, Layer> page_layers)>();

                var ignorePages = new List<string>();

                for (var i = 0; i < page_files.Count; i++)
                {
                    var page_file = page_files[i];

                    var page_xml = GetFileContent(zf, page_file);

                    var shapes = _parse_vsdx_shapes(page_xml, null, masters);

                    var page_root = XDocument.Parse(page_xml).Root;
                    var connects = _parse_connects(page_root);
                    var page_layers = _parse_layers(page_root);

                    var page_rels = _parse_rels(zf, page_file);

                    page_cache.Add(i, (shapes, connects, page_rels, page_layers));
                }

                for (var i = 0; i < page_files.Count; i++)
                {
                    var page_file = page_files[i];

                    if (!page_cache.ContainsKey(i))
                    {
                        continue;
                    }

                    var (shapes, connects, page_rels, page_layers) = page_cache[i];

                    if (shapes == null || shapes.Count == 0)
                    {
                        continue;
                    }

                    var (page_w, page_h) = (8.5f, 11.0f);

                    if (i < all_dims.Count)
                        (page_w, page_h) = all_dims[i];
                    else
                    {
                        var page_xml = GetFileContent(zf, page_file);
                        (page_w, page_h) = _parse_page_dimensions(page_xml);
                    }

                    // Background page composition
                    List<Shape> bg_shapes = null;
                    List<Connect> bg_connects = null;

                    if (bg_map.ContainsKey(i))
                    {
                        var bg_idx = bg_map[i];

                        if (page_cache.ContainsKey(bg_idx))
                        {
                            var cache = page_cache[bg_idx];

                            bg_shapes = cache.shapes;
                            bg_connects = cache.connects;
                        }
                    }

                    // Merge master_rels into page_rels for image resolution
                    var all_rels = master_rels.ToDictionary();

                    if (page_rels != null)
                    {
                        foreach (var kp in page_rels)
                        {
                            var key = kp.Key;

                            if (all_rels.ContainsKey(key))
                            {
                                all_rels[key] = page_rels[key];
                            }
                            else
                            {
                                all_rels.Add(key, kp.Value);
                            }
                        }
                    }

                    var svg_content = _shapes_to_svg(
                                    shapes, page_w, page_h, doc, masters, connects,
                                    media, all_rels, bg_shapes, bg_connects,
                                    theme_colors, page_layers);

                    svgs.Add(svg_content);
                }
            }

            return svgs;
        }

        private static SharpCompress.Readers.ReaderOptions GetZipOptions()
        {
            var cultureInfo = System.Globalization.CultureInfo.CurrentCulture;

            SharpCompress.Readers.ReaderOptions options = new SharpCompress.Readers.ReaderOptions();

            Encoding defaultEncoding = Encoding.UTF8;

            if (cultureInfo.Name == "zh-CN")
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                defaultEncoding = Encoding.GetEncoding("gbk");
            }

            options.ArchiveEncoding.Default = defaultEncoding;

            return options;
        }

        #endregion
    }
}