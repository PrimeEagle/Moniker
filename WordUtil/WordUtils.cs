using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace WordSplitter
{
    public static class WordUtils
    {
        public static void Split(string fixedSourceFile, string sourceFile, string outputDirectory, string styleName)
        {
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

                // Save last section
                if (start != null && counter > 0)
                {
                    string newFileName = $"{outputDirectory}\\{(counter).ToString().PadLeft(3, '0')}-{file}-{section}.docx";
                    SaveSection(originalDoc, newFileName, start, null);
                }
            }
        }

        private static void SaveSection(WordprocessingDocument originalDoc, string newFileName, Paragraph start, Paragraph end)
        {
            originalDoc.Save();  // Save the original document before performing the copy operations
            Thread.Sleep(5000);
            using (WordprocessingDocument newDoc = WordprocessingDocument.Create(newFileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = newDoc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Copy all namespace declarations from the original document to the new document.
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

                // Copy all other parts from original document to new document
                CopyAllParts(originalDoc.MainDocumentPart, mainPart);

                mainPart.Document.Save();
                newDoc.Save();
                newDoc.Close();
            }

            Thread.Sleep(1500);

            var tempFileName = newFileName + ".tmp.docx";
            if(File.Exists(tempFileName))
            {
                File.Delete(tempFileName);
            }
            File.Move(newFileName, tempFileName);
            ResaveDocument(tempFileName, newFileName);
            Thread.Sleep(3000);
            File.Delete(tempFileName);
        }

        public static void CopyAllParts(OpenXmlPart originalPart, OpenXmlPart newPart)
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
                OpenXmlPart newChildPart = newPart.AddPart(originalChildPart, partPair.RelationshipId); // Copy relationship ID
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
            // Create a new Word application.
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = false;  // Set to false to make Word run in the background

            // Open the document.
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(sourceFile);

            // Save the document. This can help fix minor issues.
            object fileFormat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
            doc.SaveAs2(targetFile, fileFormat);

            // Close the document and quit Word.
            doc.Close();
            wordApp.Quit();
        }

        private static bool ValidateDocument(string filename)
        {
            // Initialize the OpenXML Validator
            OpenXmlValidator validator = new OpenXmlValidator();
            int count = 0;

            // Open the Word document.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filename, true))
            {
                // Validate the Word document and capture the validation errors.
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

        public static void ReplaceFontWithStyle(string sourceFile, string font, string style, string styleFont, string styleColor, double? styleSize, bool styleBold, bool styleItalic)
        {
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
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(inputFile, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                // Collect all tables first
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

                    // Create the new paragraph
                    Paragraph newPara = new Paragraph(new Run(new Text(tableText.ToString())));

                    // Insert the new paragraph after the table's parent
                    tbl.Parent.InsertAfter(newPara, tbl);

                    // Remove the original table
                    tbl.Remove();
                }

                wordDoc.Save();
                Thread.Sleep(3000);
            }
        }

        public static void UpdateCrossReferences(string inputFile)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(inputFile, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                // Collect all tables first
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

                    // Create the new paragraph
                    Paragraph newPara = new Paragraph(new Run(new Text(tableText.ToString())));

                    // Insert the new paragraph after the table's parent
                    tbl.Parent.InsertAfter(newPara, tbl);

                    // Remove the original table
                    tbl.Remove();
                }

                wordDoc.Save();
                Thread.Sleep(3000);
            }
        }
    }
}