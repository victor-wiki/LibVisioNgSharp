namespace LibVisioSharp.Helper
{
    public class FileHelper
    {
        public static readonly Dictionary<string, string> MimeMappings = new Dictionary<string, string>()
        {
            {"png", "image/png" },
            {"jpg", "image/jpeg"},
            {"jpeg", "image/jpeg"},
            {"gif", "image/gif"},
            {"svg", "image/svg+xml"},
            {"bmp", "image/bmp"},
            {"tiff", "image/tiff"},
            {"tif", "image/tiff"},
            {"emf", "image/x-emf"},
            {"wmf", "image/x-wmf"},
            {"webp", "image/webp"},
            {"mp4", "video/mp4"},
            {"m4v", "video/mp4"},
            {"webm", "video/webm"},
            {"avi", "video/x-msvideo"},
            {"mp3", "audio/mpeg"},
            {"wav", "audio/wav"},
            {"m4a", "audio/mp4"},
            {"ogg", "audio/ogg"},
        };
    }
}
