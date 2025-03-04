using DevExpress.Portable.Input;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using System.Diagnostics;
using System.Drawing;

namespace RichEditTOCGeneration {
    class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>

        static void Main(string[] args)
        {
            Console.WriteLine("Select an approach to generate a TOC:\r\nBased on Styles - enter 1\r\nBased on outline levels - enter 2\r\nBased on TC fields - enter 3\r\n");
            var answer = Console.ReadLine();
            string documentName = "";
            switch (answer)
            {
                case "1": documentName = ApplyStyles(); break;
                case "2": documentName = AssignOutlineLevels(); break;
                case "3": documentName = AddTCFields(); break;
            }

            var p = new Process();
            p.StartInfo = new ProcessStartInfo(documentName)
            {
                UseShellExecute = true
            };
            p.Start();
        }

        private static string ApplyStyles()
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None;
                wordProcessor.LoadDocument("Employees.rtf");
                Document document = wordProcessor.Document;
                document.BeginUpdate();
                SearchForTOCEntries(document, delegate (DocumentPosition location, int level) {
                    document.Paragraphs.Get(location).Style = GetStyleForLevel(document, level);
                });
                InsertTOC(document, "\\h", true);
                document.EndUpdate();
                string documentName = "Employees_with_Styles_TOC.docx";
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml);
                Console.WriteLine(documentName+" is created");
                return documentName;
            }
        }

        private static string AssignOutlineLevels()
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None;
                wordProcessor.LoadDocument("Employees.rtf");
                Document document = wordProcessor.Document;
                document.BeginUpdate();

                SearchForTOCEntries(document, delegate (DocumentPosition location, int level) {
                    document.Paragraphs.Get(location).OutlineLevel = level;
                });
                InsertTOC(document, "\\h \\u", true);
                document.EndUpdate();
                string documentName = "Employees_with_Outlines_TOC.docx";
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml);
                Console.WriteLine(documentName + " is created");
                return documentName;
            }
        }

        private static string AddTCFields()
        {
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.Options.Hyperlinks.ModifierKeys = PortableKeys.None;
                wordProcessor.LoadDocument("Employees.rtf");
                Document document = wordProcessor.Document;
                document.BeginUpdate();
                SearchForTOCEntries(document, delegate (DocumentPosition location, int level) {
                    document.Fields.Create(location, string.Format("TC \"{0}\" \\f {1} \\l {2}",
                        document.GetText(document.Paragraphs.Get(location).Range), "defaultGroup", level));
                });
                InsertTOC(document, "\\h \\f defaultGroup", true);
                document.Fields.Update();
                document.EndUpdate();
                string documentName = "Employees_with_TCFields_TOC.docx";
                wordProcessor.SaveDocument(documentName, DocumentFormat.OpenXml);
                Console.WriteLine(documentName + " is created");
                return documentName;
            }
        }

        private static void SearchForTOCEntries(Document document, Action<DocumentPosition, int> callback)
        {
            for (int i = 0; i < document.Paragraphs.Count; i++)
            {
                DocumentRange range = document.CreateRange(document.Paragraphs[i].Range.Start, 1);
                CharacterProperties cp = document.BeginUpdateCharacters(range);
                int level = 0;

                if (cp.FontSize.Equals(14f))
                    level = 1;
                if (cp.FontSize.Equals(13f))
                    level = 2;
                if (cp.FontSize.Equals(11f))
                    level = 3;

                document.EndUpdateCharacters(cp);

                if (level != 0)
                    callback(range.Start, level);
            }
        }

        private static void InsertTOC(Document document, string switches, bool insertHeading)
        {
            if (insertHeading)
                InsertContentHeading(document);

            Field field = document.Fields.Create(document.Paragraphs[(insertHeading ? 1 : 0)].Range.Start, "TOC " + switches);
            CharacterProperties cp = document.BeginUpdateCharacters(field.Range);
            cp.Bold = false;
            cp.FontSize = 12;
            cp.ForeColor = Color.Blue;
            document.EndUpdateCharacters(cp);
            document.InsertSection(field.Range.End);
            field.Update();
        }

        private static void InsertContentHeading(Document document)
        {
            DocumentRange range = document.InsertText(document.Range.Start, "Contents\r\n");
            CharacterProperties cp = document.BeginUpdateCharacters(range);
            cp.FontSize = 18;
            cp.ForeColor = Color.DarkBlue;
            document.EndUpdateCharacters(cp);
            Paragraph paragraph = document.Paragraphs[0];
            paragraph.Alignment = ParagraphAlignment.Center;
            paragraph.Style = document.ParagraphStyles["Normal"];
            paragraph.OutlineLevel = 0;
        }

        private static ParagraphStyle GetStyleForLevel(Document document, int level)
        {
            string styleName = "Paragraph Level " + level.ToString();
            ParagraphStyle paragraphStyle = document.ParagraphStyles[styleName];

            if (paragraphStyle == null)
            {
                paragraphStyle = document.ParagraphStyles.CreateNew();
                paragraphStyle.Name = styleName;
                paragraphStyle.Parent = document.ParagraphStyles["Normal"];
                paragraphStyle.OutlineLevel = level;
                document.ParagraphStyles.Add(paragraphStyle);
            }

            return paragraphStyle;
        }
    }
}