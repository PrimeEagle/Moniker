namespace WordUtil
{
    public class Parameters
    {
        public string SourceFile { get; set; }
        public string FixedSourceFile { get; set; }
        public string OutputDirectory { get; set; }
        public string SplitStyleName { get; set; }
        public string ReplaceFont { get; set; }
        public string Style { get; set; }
        public string StyleFont { get; set; }
        public string StyleColor { get; set; }
        public double? StyleSize { get; set; }
        public bool StyleBold { get; set; }
        public bool StyleItalic { get; set; }
        public string MappingFile { get; set; }
        public string TableCaptionStyleName { get; set; }
        public string FigureCaptionStyleName { get; set; }
        public string XRefStyleName { get; set; }
    }
}