namespace WordSplitter
{
    public class Parameters
    {
        public string SourceFile { get; set; }
        public string FixedSourceFile { get; set; }
        public string OutputDirectory { get; set; }
        public string SplitStyleName { get; set; }
        public string ReplaceFont { get; set; }
        public string ReplacementStyle { get; set; }
        public string ReplacementStyleFont { get; set; }
        public string ReplacementStyleColor { get; set; }
        public double? ReplacementStyleSize { get; set; }
        public bool ReplacementStyleBold { get; set; }
        public bool ReplacementStyleItalic { get; set; }
        public string MappingFile { get; set; }
    }
}
