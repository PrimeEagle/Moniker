using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic.FileIO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordUtil
{
    public static class WordUtils
    {
        private static Dictionary<string, Paragraph> uniqueTagMappings = new Dictionary<string, Paragraph>();
        private static int bookmarkId = 1;

        public static void Split(string fixedSourceFile, string sourceFile, string outputDirectory, string styleName)
        {
            Console.WriteLine($"splitting document on style '{styleName}'...");

            string file = Path.GetFileNameWithoutExtension(sourceFile);

            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            using (WordprocessingDocument originalDoc = WordprocessingDocument.Open(fixedSourceFile, false))
            {
                Body originalBody = originalDoc.MainDocumentPart.Document.Body;
                Paragraph start = originalBody.Elements<Paragraph>().FirstOrDefault();
                string section = "";
                string lastSection = "";
                int counter = 0;

                foreach (var para in originalBody.Elements<Paragraph>())
                {
                    var style = para.ParagraphProperties?.ParagraphStyleId?.Val;

                    if (style != null && style.Value.Contains(styleName))
                    {
                        section = para.InnerText.Trim();

                        if (!string.IsNullOrWhiteSpace(section))
                        {
                            counter++;
                            if (counter >= 2)
                            {
                                string newFileName = $"{outputDirectory}\\{(counter - 1).ToString().PadLeft(3, '0')}-{file}-{lastSection}.docx";
                                SaveSection(originalDoc, newFileName, start, para);
                                start = para;
                            }

                            lastSection = section;
                        }
                    }
                }

                if (start != null && counter > 0)
                {
                    string newFileName = $"{outputDirectory}\\{(counter).ToString().PadLeft(3, '0')}-{file}-{section}.docx";
                    SaveSection(originalDoc, newFileName, start, null);
                }
            }
        }

        private static void SaveSection(WordprocessingDocument originalDoc, string newFileName, Paragraph start, Paragraph end)
        {
            originalDoc.Save();
            Thread.Sleep(5000);
            using (WordprocessingDocument newDoc = WordprocessingDocument.Create(newFileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = newDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                foreach (var ns in originalDoc.MainDocumentPart.Document.NamespaceDeclarations)
                {
                    mainPart.Document.AddNamespaceDeclaration(ns.Key, ns.Value);
                }

                bool copy = false;

                foreach (var elem in originalDoc.MainDocumentPart.Document.Body.Elements())
                {
                    if (elem == start)
                    {
                        copy = true;
                    }

                    if (elem == end)
                    {
                        copy = false;
                    }

                    if (copy)
                    {
                        var clonedElement = (OpenXmlElement)elem.CloneNode(true);
                        mainPart.Document.Body.Append(clonedElement);
                    }
                }

                CopyAllParts(originalDoc.MainDocumentPart, mainPart);

                mainPart.Document.Save();
                newDoc.Save();
                newDoc.Close();
            }

            Thread.Sleep(1500);

            var tempFileName = newFileName + ".tmp.docx";
            if (File.Exists(tempFileName))
            {
                File.Delete(tempFileName);
            }
            File.Move(newFileName, tempFileName);
            ResaveDocument(tempFileName, newFileName);
            Thread.Sleep(3000);
            File.Delete(tempFileName);
        }

        private static void CopyAllParts(OpenXmlPart originalPart, OpenXmlPart newPart)
        {
            using (Stream stream = originalPart.GetStream(FileMode.Open))
            using (MemoryStream memStream = new MemoryStream())
            using (Stream newStream = newPart.GetStream(FileMode.OpenOrCreate))
            {
                if (stream.Length > 0)
                {
                    stream.CopyTo(memStream);
                    memStream.Position = 0;
                    memStream.CopyTo(newStream);
                }
            }

            foreach (IdPartPair partPair in originalPart.Parts)
            {
                OpenXmlPart originalChildPart = partPair.OpenXmlPart;
                OpenXmlPart newChildPart = newPart.AddPart(originalChildPart, partPair.RelationshipId);
                CopyAllParts(originalChildPart, newChildPart);
            }

            foreach (ExternalRelationship extRel in originalPart.ExternalRelationships)
            {
                newPart.AddExternalRelationship(extRel.RelationshipType, extRel.Uri, extRel.Id);
            }

            foreach (HyperlinkRelationship hyperlink in originalPart.HyperlinkRelationships)
            {
                newPart.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, hyperlink.Id);
            }
        }

        public static void ResaveDocument(string sourceFile, string targetFile)
        {
            Console.WriteLine($"re-saving document '{Path.GetFileName(sourceFile)}'...");

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(sourceFile);
            object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
            doc.SaveAs2(targetFile, fileFormat);
            doc.Close();
            wordApp.Quit();
        }

        public static bool ValidateDocument(string filename)
        {
            Console.WriteLine($"validating document...");

            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                foreach (ValidationErrorInfo error in validator.Validate(wordDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                if (count == 0)
                {
                    Console.WriteLine("The document is valid.");
                }
                else
                {
                    Console.WriteLine($"Total Errors: {count}");
                }
            }
            return count == 0;
        }

        public static void CreateParagraphStyle(StyleDefinitionsPart stylesPart, string style, string styleFont = null, string styleColor = null, double? styleSize = null, bool styleBold = false, bool styleItalic = false)
        {
            Console.WriteLine($"creating paragraph style '{style}'...");

            string styleId = style;
            Style existingStyle = stylesPart.Styles.Elements<Style>()
                                    .FirstOrDefault(s => s.StyleId != null && s.StyleId.Value.Equals(styleId));

            if (existingStyle == null)
            {
                Style newStyle = new Style()
                {
                    StyleId = styleId,
                    StyleName = new StyleName() { Val = styleId },
                };

                if (!string.IsNullOrEmpty(styleFont) || styleSize.HasValue || !string.IsNullOrEmpty(styleColor) || styleBold || styleItalic)
                {
                    newStyle.StyleRunProperties = new StyleRunProperties();

                    if (!string.IsNullOrEmpty(styleFont))
                    {
                        newStyle.StyleRunProperties.RunFonts = new RunFonts();
                        newStyle.StyleRunProperties.RunFonts.Ascii = styleFont;
                    }

                    if (styleSize.HasValue)
                    {
                        newStyle.StyleRunProperties.FontSize = new FontSize();
                        newStyle.StyleRunProperties.FontSize.Val = (styleSize.Value * 2.0d).ToString();
                    }

                    if (!string.IsNullOrEmpty(styleColor))
                    {
                        newStyle.StyleRunProperties.Color = new Color();
                        newStyle.StyleRunProperties.Color.Val = styleColor;
                    }

                    if (styleBold)
                    {
                        newStyle.StyleRunProperties.Bold = new Bold();
                    }

                    if (styleItalic)
                    {
                        newStyle.StyleRunProperties.Italic = new Italic();
                    }
                }

                stylesPart.Styles.AppendChild(newStyle);
                stylesPart.Styles.Save();
            }
        }

        public static void CreateCharacterStyle(StyleDefinitionsPart stylesPart, string style, string styleFont = null, string styleColor = null, double? styleSize = null, bool styleBold = false, bool styleItalic = false)
        {
            Console.WriteLine($"creating character style '{style}'...");

            string styleId = style;
            Style existingStyle = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId != null && s.StyleId.Value.Equals(styleId));

            if (existingStyle == null)
            {
                Style newStyle = new Style()
                {
                    StyleId = styleId,
                    StyleName = new StyleName() { Val = styleId },
                    Type = StyleValues.Character
                };

                if (!string.IsNullOrEmpty(styleFont) || styleSize.HasValue || !string.IsNullOrEmpty(styleColor) || styleBold || styleItalic)
                {
                    newStyle.StyleRunProperties = new StyleRunProperties();

                    if (!string.IsNullOrEmpty(styleFont))
                    {
                        newStyle.StyleRunProperties.RunFonts = new RunFonts();
                        newStyle.StyleRunProperties.RunFonts.Ascii = styleFont;
                    }

                    if (styleSize.HasValue)
                    {
                        newStyle.StyleRunProperties.FontSize = new FontSize();
                        newStyle.StyleRunProperties.FontSize.Val = (styleSize.Value * 2.0d).ToString();
                    }

                    if (!string.IsNullOrEmpty(styleColor))
                    {
                        newStyle.StyleRunProperties.Color = new Color();
                        newStyle.StyleRunProperties.Color.Val = styleColor;
                    }

                    if (styleBold)
                    {
                        newStyle.StyleRunProperties.Bold = new Bold();
                    }

                    if (styleItalic)
                    {
                        newStyle.StyleRunProperties.Italic = new Italic();
                    }
                }

                stylesPart.Styles.AppendChild(newStyle);
                stylesPart.Styles.Save();
            }
        }

        public static void ReplaceFontWithStyle(string sourceFile, string font, string style)
        {
            Console.WriteLine($"replacing font '{font}' with style '{style}'...");

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sourceFile, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                StyleDefinitionsPart stylesPart;

                if (mainPart.StyleDefinitionsPart == null)
                {
                    stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();
                }
                else
                {
                    stylesPart = mainPart.StyleDefinitionsPart;
                }

                WordUtils.GetOrCreateParagraphStyle(stylesPart, style);

                foreach (Paragraph para in mainPart.Document.Body.OfType<Paragraph>())
                {
                    bool allRunsHaveUniqueFont = true;
                    foreach (Run run in para.Elements<Run>())
                    {
                        if (run.RunProperties != null && run.RunProperties.RunFonts != null)
                        {
                            RunFonts runFont = run.RunProperties.RunFonts;
                            if (runFont.Ascii != font)
                            {
                                allRunsHaveUniqueFont = false;
                                break;
                            }
                        }
                    }

                    if (allRunsHaveUniqueFont)
                    {
                        para.ParagraphProperties = new ParagraphProperties(
                            new ParagraphStyleId() { Val = style });
                    }
                }

                mainPart.Document.Save();
                Thread.Sleep(3000);
            }
        }

        public static void ReplaceTables(string inputFile)
        {
            Console.WriteLine($"replacing tables with custom tags...");

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(inputFile, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                List<Table> tables = new List<Table>(body.Descendants<Table>());

                foreach (var tbl in tables)
                {
                    StringBuilder tableText = new StringBuilder("START_TABLE|");

                    foreach (var row in tbl.Descendants<TableRow>())
                    {
                        foreach (var cell in row.Descendants<TableCell>())
                        {
                            tableText.AppendFormat("{0}^", cell.InnerText);
                        }
                        tableText.Length--;  // Removes the last comma
                        tableText.Append("|");
                    }
                    tableText.Length--;  // Removes the last |
                    tableText.Append("|END_TABLE");
                    Paragraph newPara = new Paragraph(new Run(new Text(tableText.ToString())));

                    tbl.Parent.InsertAfter(newPara, tbl);
                    tbl.Remove();
                }

                wordDoc.Save();
                Thread.Sleep(3000);
            }
        }

        public static void ProcessCrossReferences(string filePath, string tableCaptionStyleName, string figureCaptionStyleName, string xRefStyleName)
        {
            Console.WriteLine($"processing cross-references with table style '{tableCaptionStyleName}', figure style '{figureCaptionStyleName}', and cross-reference style '{xRefStyleName}'...");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                var body = doc.MainDocumentPart.Document.Body;

                ProcessContentTags(body, "Table", tableCaptionStyleName); 
                ProcessContentTags(body, "Figure", figureCaptionStyleName);
                ProcessXRefs(doc, xRefStyleName, ":", ".");

                doc.MainDocumentPart.Document.Save();
            }
        }

        private static void ProcessContentTags(Body body, string contentType, string styleName)
        {
            string tagPattern = $@"\[{contentType}:(.*?):(.*?)(?::(.*?))?\]";

            foreach (var para in body.Descendants<Paragraph>())
            {
                var match = Regex.Match(para.InnerText, tagPattern);
                if (match.Success)
                {
                    string uniqueTag = match.Groups[1].Value;
                    string name = match.Groups[2].Value;
                    string description = match.Groups[3].Value;

                    para.RemoveAllChildren<Run>();

                    string replacedContent = $"{contentType} ###: {name}";
                    para.AppendChild(new Run(new Text(replacedContent)));

                    if (!string.IsNullOrEmpty(styleName))
                    {
                        para.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = styleName });
                    }

                    var info = $"{contentType} ###:{name}:{description}";
                    uniqueTagMappings.Add(uniqueTag, para);
                }
            }
        }

        private static void ProcessXRefs(WordprocessingDocument doc, string xRefStyleName, string tableDelimiter = null, string figureDelimiter = null)
        {
            var body = doc.MainDocumentPart.Document.Body;

            foreach (var para in body.Descendants<Paragraph>())
            {
                if (para.InnerText.Contains("[XRef:"))
                {
                    string paraText = para.InnerText;
                    var match = Regex.Match(paraText, @"\[XRef:(.*?):(.*?)\]");
                    if (match.Success)
                    {
                        var uniqueTag = match.Groups[1].Value;
                        var options = match.Groups[2].Value.Split(',');

                        if (uniqueTagMappings.TryGetValue(uniqueTag, out var refPara))
                        {
                            var refText = refPara.InnerText;
                            var refParts = refText.Split(":");
                            var label = refParts[0];
                            var name = refParts[1];
                            string desc = null;
                            if (refParts.Count() > 2)
                            {
                                desc = refParts[2];
                            }

                            string output = label;
                            var delimiter = ":";
                            if (label.ToLower().Contains("table") && !string.IsNullOrEmpty(tableDelimiter)) { delimiter = tableDelimiter; }
                            if (label.ToLower().Contains("figure") && !string.IsNullOrEmpty(figureDelimiter)) { delimiter = figureDelimiter; }

                            if (options.Contains("name")) { output += $"{delimiter} {name}"; }
                            if (options.Contains("desc") && !options.Contains("name")) { output += $"{delimiter} {desc}"; }
                            if (options.Contains("desc") && options.Contains("name")) { output += $" ({desc})"; }
                            if (options.Contains("page")) { output += $" on page #"; }

                            AddBookmark(refPara, $"B-{uniqueTag}", bookmarkId++);

                            foreach (var run in para.Descendants<Run>())
                            {
                                if (run.InnerText.Contains(match.Value))
                                {
                                    run.GetFirstChild<Text>().Text = run.GetFirstChild<Text>().Text.Replace(match.Value, string.Empty);
                                    break;
                                }
                            }

                            AddCrossReference(para, output, $"B-{uniqueTag}");

                            if (!string.IsNullOrEmpty(xRefStyleName))
                            {
                                para.Descendants<Run>().First().RunProperties = new RunProperties(new RunStyle() { Val = xRefStyleName });
                            }
                        }
                    }
                }
            }
        }

        private static void AddBookmark(Paragraph paragraph, string name, int id)
        {
            bool bookmarkExists = paragraph.Descendants<BookmarkStart>().Any(b => b.Name == name);

            if (!bookmarkExists)
            {
                var bookmarkStart = new BookmarkStart() { Name = name, Id = id.ToString() };
                var bookmarkEnd = new BookmarkEnd() { Id = id.ToString() };

                var paraProps = paragraph.GetFirstChild<ParagraphProperties>();
                if (paraProps != null)
                {
                    paraProps.InsertAfterSelf(bookmarkStart);
                }
                else
                {
                    paragraph.InsertAt(bookmarkStart, 0);
                }

                bookmarkStart.InsertAfterSelf(bookmarkEnd);
            }
        }


        private static void AddCrossReference(Paragraph paragraph, string text, string destinationBookmarkName)
        {
             var fieldBegin = new FieldChar
            {
                FieldCharType = FieldCharValues.Begin
            };

            var fieldCode = new Run(new FieldCode
            {
                Text = $" REF {destinationBookmarkName} \\h "
            });

            var fieldSeparator = new FieldChar
            {
                FieldCharType = FieldCharValues.Separate
            };

            var fieldText = new Run(new Text
            {
                Text = text
            });

            var fieldEnd = new FieldChar
            {
                FieldCharType = FieldCharValues.End
            };

            var run = new Run();
            run.Append(fieldBegin, fieldCode, fieldSeparator, fieldText, fieldEnd);
            paragraph.Append(run);
        }

        private static Style GetOrCreateParagraphStyle(StyleDefinitionsPart stylesPart, string styleName)
        {
            Style style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleName);

            if (style == null)
            {
                CreateParagraphStyle(stylesPart, styleName);

                stylesPart.Styles.Append(style);
                stylesPart.Styles.Save();

                style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleName);
            }

            return style;
        }

        private static Style GetOrCreateCharacterStyle(StyleDefinitionsPart stylesPart, string styleName)
        {
            Style style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleName);

            if (style == null)
            {
                CreateCharacterStyle(stylesPart, styleName);

                stylesPart.Styles.Append(style);
                stylesPart.Styles.Save();

                style = stylesPart.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId == styleName);
            }

            return style;
        }

        private static void ReplaceParagraphText(Paragraph para, string newText)
        {
            para.RemoveAllChildren<Run>();
            para.AppendChild(new Run(new Text(newText)));
        }

        private static void ApplyParagraphStyleToParagraph(WordprocessingDocument doc, string styleId, Paragraph para)
        {
            para.RemoveAllChildren<ParagraphStyleId>();
            para.AppendChild(new ParagraphStyleId() { Val = styleId });
        }

        private static void ApplyCharacterStyleToRun(string styleId, Run run)
        {
            run.RemoveAllChildren<RunStyle>();
            run.AppendChild(new RunStyle() { Val = styleId });
        }
    }
}