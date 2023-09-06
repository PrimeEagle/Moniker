using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;

namespace WordSplitter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 7 && args[0] == "split" && args.Contains("-i") && args.Contains("-o") && args.Contains("-s"))
            {
                var param = PerformCommon(args);

                WordUtils.Split(param.FixedSourceFile, param.SourceFile, param.OutputDirectory, param.SplitStyleName);
            }
            else if (args.Length >= 7 && args[0] == "rfs" && args.Contains("-i") && args.Contains("-f") && args.Contains("-s"))
            {
                var param = PerformCommon(args);

                WordUtils.ReplaceFontWithStyle(param.FixedSourceFile, param.ReplaceFont, param.ReplacementStyle, param.ReplacementStyleFont, param.ReplacementStyleColor, param.ReplacementStyleSize, param.ReplacementStyleBold, param.ReplacementStyleItalic);
            }
            else if (args.Length == 5 && args[0] == "rfs" && args.Contains("-i") && args.Contains("-m"))
            {
                var param = PerformCommon(args);

                using (var sr = new StreamReader(param.MappingFile))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        var split = line.Split(',');
                        var font = GetCsvValue(split[0]);
                        var style = GetCsvValue(split[1]);
                        string styleFont = split.Count() > 2 ? GetCsvValue(split[2]) : null;
                        string styleColor = split.Count() > 3 ? GetCsvValue(split[3]) : null;
                        double? styleSize = split.Count() > 4 ? double.Parse(GetCsvValue(split[4])) : (double?)null;
                        bool styleBold = split.Count() > 5 ? bool.Parse(GetCsvValue(split[5])) : false;
                        bool styleItalic = split.Count() > 6 ? bool.Parse(GetCsvValue(split[6])) : false;

                        WordUtils.ReplaceFontWithStyle(param.FixedSourceFile, param.ReplaceFont, param.ReplacementStyle, param.ReplacementStyleFont, param.ReplacementStyleColor, param.ReplacementStyleSize, param.ReplacementStyleBold, param.ReplacementStyleItalic);
                    }
                }
            }
            else if (args.Length == 3 && args[0] == "ttg" && args.Contains("-i"))
            {
                var param = PerformCommon(args);

                WordUtils.ReplaceTables(param.FixedSourceFile);
            }
            else if (args.Length == 3 && args[0] == "cref" && args.Contains("-i"))
            {
                var param = PerformCommon(args);

                WordUtils.UpdateCrossReferences(param.FixedSourceFile);
            }
            else if(args.Length == 9 && args[0] == "batch" && args.Contains("-i") && args.Contains("-o") && args.Contains("-s") && args.Contains("-m"))
            {
                var param = PerformCommon(args);

                using (var sr = new StreamReader(param.MappingFile))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        var split = line.Split(',');
                        var font = GetCsvValue(split[0]);
                        var style = GetCsvValue(split[1]);
                        string styleFont = split.Count() > 2 ? GetCsvValue(split[2]) : null;
                        string styleColor = split.Count() > 3 ? GetCsvValue(split[3]) : null;
                        double? styleSize = split.Count() > 4 ? double.Parse(GetCsvValue(split[4])) : (double?)null;
                        bool styleBold = split.Count() > 5 ? bool.Parse(GetCsvValue(split[5])) : false;
                        bool styleItalic = split.Count() > 6 ? bool.Parse(GetCsvValue(split[6])) : false;

                        WordUtils.ReplaceFontWithStyle(param.FixedSourceFile, font, style, styleFont, styleColor, styleSize, styleBold, styleItalic);
                    }
                }
                Thread.Sleep(13000);
                WordUtils.ReplaceTables(param.FixedSourceFile);
                WordUtils.UpdateCrossReferences(param.FixedSourceFile);
                WordUtils.Split(param.FixedSourceFile, param.SourceFile, param.OutputDirectory, param.SplitStyleName);
            }
            else
            {
                Console.WriteLine("usage for batch mode:");
                Console.WriteLine("  WordSplitter.exe batch -i [file] -o [directory] -s [style] -m [mapping file]");
                Console.WriteLine("    -i - input file");
                Console.WriteLine("    -o - output directory");
                Console.WriteLine("    -s - style name to match for splitting");
                Console.WriteLine("    -m - font-style mapping file (comma delimited)");
                Console.WriteLine();
                Console.WriteLine("usage for splitting a Word document:");
                Console.WriteLine("  WordSplitter.exe split -i [file] -o [directory] -s [style]");
                Console.WriteLine("    -i - input file");
                Console.WriteLine("    -o - output directory");
                Console.WriteLine("    -s - style name to match for splitting");
                Console.WriteLine();
                Console.WriteLine("usage for replacing a single font with a style in a Word document:");
                Console.WriteLine("  WordSplitter.exe rfs -i [file] -f [font] -s [style] {-sf [style font]} {-ss [style font size]} {-sb} {-si} {-sc [style color]");
                Console.WriteLine("    -i - input file");
                Console.WriteLine("    -f - font name");
                Console.WriteLine("    -s - style name");
                Console.WriteLine("    -sf  style font name (if style doesn't exist) - OPTIONAL");
                Console.WriteLine("    -ss  style font size (if style doesn't exist) - OPTIONAL"); 
                Console.WriteLine("    -sb  style bold (if style doesn't exist) - OPTIONAL");
                Console.WriteLine("    -si  style italic (if style doesn't exist) - OPTIONAL");
                Console.WriteLine("    -sc  style font color (if style doesn't exist) - OPTIONAL");
                Console.WriteLine();
                Console.WriteLine("usage for replacing multiple fonts with styles in a Word document:");
                Console.WriteLine("  WordSplitter.exe rfs -i [file] -m [mapping file]");
                Console.WriteLine("    -i - input file");
                Console.WriteLine("    -m - font-style mapping file (comma delimited)");
                Console.WriteLine();
                Console.WriteLine("usage for replacing tables with tags in a Word document:");
                Console.WriteLine("  WordSplitter.exe ttg -i [file]");
                Console.WriteLine("    -i - input file");

                Environment.Exit(0);
            }
        }

        private static string GetCsvValue(string value)
        {
            string result = value;

            if (result.StartsWith("\"")) { result = result.Substring(1); }
            if (result.EndsWith("\"")) { result = result.Substring(0, result.Length - 1); }

            return result;
        }
        
        private static Parameters PerformCommon(string[] args)
        {
            var param = new Parameters();

            param.SourceFile = args.Contains("-i") ? GetCsvValue(args.SkipWhile(a => a != "-i").ElementAt(1)) : null;
            param.OutputDirectory = args.Contains("-o") ? GetCsvValue(args.SkipWhile(a => a != "-o").ElementAt(1)) : null;
            param.SplitStyleName = args.Contains("-s") ? GetCsvValue(args.SkipWhile(a => a != "-s").ElementAt(1)) : null;
            param.ReplaceFont = args.Contains("-f") ? GetCsvValue(args.SkipWhile(a => a != "-f").ElementAt(1)) : null;
            param.ReplacementStyle = args.Contains("-s") ? GetCsvValue(args.SkipWhile(a => a != "-s").ElementAt(1)) : null;
            param.ReplacementStyleFont = args.Contains("-sf") ? args.SkipWhile(a => a != "-sf").Any() ? args.SkipWhile(a => a != "-sf").ElementAt(1) : null : null;
            param.ReplacementStyleColor = args.Contains("-sc") ? args.SkipWhile(a => a != "-sc").Any() ? args.SkipWhile(a => a != "-sc").ElementAt(1) : null : null;
            param.ReplacementStyleSize = args.Contains("-ss") ? args.SkipWhile(a => a != "-ss").Any() ? double.Parse(args.SkipWhile(a => a != "-ss").ElementAt(1)) : (double?)null : null;
            param.ReplacementStyleBold = args.Contains("-sb") ? args.SkipWhile(a => a != "-sb").Any() : false;
            param.ReplacementStyleItalic = args.Contains("-si") ? args.SkipWhile(a => a != "-si").Any() : false;
            param.MappingFile = args.Contains("-m") ? GetCsvValue(args.SkipWhile(a => a != "-m").ElementAt(1)) : null;

            if (!string.IsNullOrEmpty(param.SourceFile) && !File.Exists(param.SourceFile))
            {
                Console.WriteLine($"Input file '{param.SourceFile}' not found.");
                Environment.Exit(0);
            }

            if (!string.IsNullOrEmpty(param.MappingFile) && !File.Exists(param.MappingFile))
            {
                Console.WriteLine($"Mapping file '{param.MappingFile}' not found.");
                Environment.Exit(0);
            }

            string folder = Path.GetDirectoryName(param.SourceFile);
            string file = Path.GetFileNameWithoutExtension(param.SourceFile);
            string ext = Path.GetExtension(param.SourceFile);
            string fixedFile = $"{folder}\\{file}-fixed{ext}";
            
            WordUtils.ResaveDocument(param.SourceFile, fixedFile);
            param.FixedSourceFile = fixedFile;

            return param;
        }
    }
}