using System.Text.RegularExpressions;

namespace LibVisioSharp.Helper
{
    public class ColorHelper
    {
        // Visio color index table (standard colors)
        public static readonly Dictionary<int, string> _VISIO_COLORS = new Dictionary<int, string>()
        {
            {0, "#000000"},  // Black
            {1, "#FFFFFF"},  // White
            {2, "#FF0000"},  // Red
            {3, "#00FF00"},  // Green
            {4, "#0000FF"},  // Blue
            {5, "#FFFF00"},  // Yellow
            {6, "#FF00FF"},  // Magenta
            {7, "#00FFFF"},  // Cyan
            {8, "#800000"},  // Dark Red
            {9, "#008000"},  // Dark Green
            {10, "#000080"}, // Dark Blue
            {11, "#808000"}, // Dark Yellow (Olive)
            {12, "#800080"}, // Dark Magenta (Purple)
            {13, "#008080"}, // Dark Cyan (Teal)
            {14, "#C0C0C0"}, // Light Gray
            {15, "#808080"}, // Dark Gray
            {16, "#993366"}, // Rose
            {17, "#333399"}, // Indigo
            {18, "#333333"}, // Charcoal
            {19, "#003300"}, // Forest
            {20, "#003366"}, // Marine
            {21, "#993300"}, // Brown
            {22, "#993366"}, // Plum
            {23, "#333399"}, // Navy
            {24, "#E6E6E6"} // Pale Gray
        };

        /// <summary>
        /// Lighten a hex color by blending towards white.
        /// </summary>
        /// <param name="hex_color"></param>
        /// <param name="factor">factor=0.0 returns original, factor = 1.0 returns white.</param>
        /// <returns></returns>
        public static string LightenColor(string hex_color, float factor = 0.7f)
        {
            hex_color = hex_color.Trim().TrimStart('#');

            if (hex_color.Length != 6)
            {
                return "#E8E8E8";
            }

            int r, g, b;

            try
            {
                r = Convert.ToInt16(hex_color.Substring(0, 2));
                g = Convert.ToInt16(hex_color.Substring(2, 2));
                b = Convert.ToInt16(hex_color.Substring(4, 2));
            }
            catch (Exception ex)
            {
                return "#E8E8E8";
            }

            r = (int)(r + (255 - r) * factor);
            g = (int)(g + (255 - g) * factor);
            b = (int)(b + (255 - b) * factor);

            return $"#{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>
        /// Check if a color is black or near-black.
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static bool IsBlack(string color)
        {
            if (color == null)
            {
                return false;
            }

            var c = color.Trim().ToUpper();

            return (new string[3] { "#000000", "#000", "0" }).Contains(c);
        }

        /// <summary>
        /// Check if a color is dark (luminance < 0.4).
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static bool IsDarkColor(string color)
        {
            if (color == null || color == "none")
            {
                return false;
            }

            var c = color.Trim().ToUpper();

            if (c.Length == 6)
            {
                try
                {
                    var r = Convert.ToInt16(c.Substring(0, 2));
                    var g = Convert.ToInt16(c.Substring(2, 2));
                    var b = Convert.ToInt16(c.Substring(4, 2));

                    // Relative luminance
                    var lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255.0;

                    return lum < 0.4;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }

            return false;
        }

        /// <summary>
        /// Convert Visio HSL (h=0-255, s=0-255, l=0-255) to #RRGGBB.
        /// </summary>
        /// <param name="h"></param>
        /// <param name="s"></param>
        /// <param name="l"></param>
        /// <returns></returns>
        public static string _hsl_to_rgb(int h, int s, int l)
        {
            // Normalize to 0-1 range
            var hf = (h / 255.0) * 360.0;
            var sf = s / 255.0;
            var lf = l / 255.0;

            double r, g, b;

            // HSL to RGB conversion
            if (sf == 0)
            {
                r = g = b = lf;
            }
            else
            {
                Func<double, double, double, double> hue2rgb = (p, q, t) =>
                {
                    if (t < 0) t += 1;
                    if (t > 1) t -= 1;
                    if (t < 1.0 / 6) return p + (q - p) * 6 * t;
                    if (t < 1.0 / 2) return q;
                    if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;

                    return p;
                };

                var q = lf < 0.5 ? lf * (1 + sf) : (lf + sf - lf * sf);
                var p = 2 * lf - q;
                var hn = hf / 360.0;

                r = hue2rgb(p, q, hn + 1.0 / 3);
                g = hue2rgb(p, q, hn);
                b = hue2rgb(p, q, hn - 1.0 / 3);
            }

            return $"#{((int)(r * 255)):X2}{((int)(g * 255)):X2}{((int)(b * 255)):X2}";
        }

        /// <summary>
        /// Handles: color index, #RRGGBB, RGB(r,g,b), HSL(h,s,l), THEMEVAL(), etc.
        /// </summary>
        /// <param name="val"></param>
        /// <param name="theme_colors"></param>
        /// <returns>Returns empty string for unresolvable values (caller decides default).</returns>
        public static string ResolveColor(string val, Dictionary<string, string> theme_colors)
        {
            if (val == null)
            {
                return null;
            }

            val = val.Trim();

            if (val.Contains("THEMEVAL") || val.Contains("THEMEGUARD"))
            {
                if (theme_colors != null)
                {
                    var m = Regex.Match(val, @"THEMEVAL\s*\(\s*""?(\w+)""?", RegexOptions.IgnoreCase);

                    if (m.Success)
                    {
                        var key = m.Groups[1].Value.ToLower();

                        if (theme_colors.ContainsKey(key))
                        {
                            return theme_colors[key];
                        }

                        if (int.TryParse(key, out var idx))
                        {
                            if (theme_colors.ContainsKey(idx.ToString()))
                            {
                                return theme_colors[idx.ToString()];
                            }

                        }
                    }

                    // THEMEGUARD(THEMEVAL(...))
                    var m2 = Regex.Match(val, @"THEMEVAL\s*\(\s*(\d+)", RegexOptions.IgnoreCase);
                    if (m2.Success)
                    {
                        var idx = m2.Groups[1].Value; // type: ignore[assignment]

                        if (theme_colors.ContainsKey(idx))
                        {
                            return theme_colors[idx]; // type: ignore[index]
                        }
                    }
                }

                return null;
            }

            if (val == "Inh" || val.StartsWith("=") || val.Contains("THEME"))
            {
                return null;
            }

            // #RRGGBB or #RGB
            if (val.StartsWith('#'))
            {
                return val;
            }

            var m3 = Regex.Match(val, @"HSL\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", RegexOptions.IgnoreCase);

            if(m3.Success)
            {
                return _hsl_to_rgb(int.Parse(m3.Groups[1].Value), int.Parse(m3.Groups[2].Value), int.Parse(m3.Groups[3].Value));
            }

            if(int.TryParse(val, out int _idx))
            {
                return _VISIO_COLORS.ContainsKey(_idx) ? _VISIO_COLORS[_idx] : null;
            }

            return null;        
        }
    }
}
