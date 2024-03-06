using NPOI.XWPF.UserModel;
using System.Text.RegularExpressions;

namespace Appeon.DotnetDemo.DocumentWriter
{
    public class DocxWriter
    {
        private static readonly Regex ArgumentDelimiterRegex = new(@"\{([^{}]+)\}");

        public static int OpenDocument(string path, out XWPFDocument? document, out string? error)
        {
            error = null;
            document = null;
            try
            {
                using var stream = new FileStream(path, mode: FileMode.Open, FileAccess.Read);
                document = new XWPFDocument(stream);
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;

            }
        }

        public static int CreateDocument(out XWPFDocument? document, out string? error)
        {
            error = null;
            document = null;

            try
            {
                document = new XWPFDocument();
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }

        private static Dictionary<string, string> ConvertToDictionary(string[] data, string separator, string nullString)
        {
            var dict = new Dictionary<string, string>();

            foreach (var entry in data)
            {
                var tokens = entry.Split(separator);
                dict.Add(tokens[0], tokens.Length > 1 ? tokens[1] : nullString);
            }

            return dict;
        }

        private static void ReplaceParagraphContents(XWPFParagraph paragraph, Dictionary<string, string> arguments)
        {
            if (paragraph.ParagraphText.Length > 0)
            {
                var parText = paragraph.ParagraphText;

                var matches = ArgumentDelimiterRegex.Matches(parText) as IReadOnlyList<Match>;

                foreach (var match in matches)
                {
                    if (!match.Success)
                    {
                        continue;
                    }

                    var matchString = match.Groups[0].Value;
                    var matchContents = match.Groups[1].Value;

                    if (arguments.ContainsKey(matchContents))
                    {
                        parText = parText.Replace(matchString, arguments[matchContents]);
                    }

                }

                paragraph.ReplaceText(paragraph.ParagraphText, parText);
            }
        }

        private static void ProcessBodyElements(IBodyElement bodyElement, Dictionary<string, string> arguments)
        {
            switch (bodyElement)
            {
                case XWPFParagraph paragraph:
                    {
                        ReplaceParagraphContents(paragraph, arguments);
                        break;
                    }

                case XWPFTable table:
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {
                                foreach (var element in cell.BodyElements)
                                {
                                    ProcessBodyElements(element, arguments);
                                }
                            }
                        }
                        break;
                    }
            }
        }

        private static void ReplaceStringInDocument(IBodyElement bodyElement, string search, string replaceWith)
        {
            switch (bodyElement)
            {
                case XWPFParagraph paragraph:
                    {
                        ReplaceStringInParagraph(ref paragraph, search, replaceWith);
                        break;
                    }

                case XWPFTable table:
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.GetTableCells())
                            {
                                foreach (var element in cell.BodyElements)
                                {
                                    ReplaceStringInDocument(element, search, replaceWith);
                                }
                            }
                        }
                        break;
                    }
            }
        }

        private static void ReplaceStringInParagraph(ref XWPFParagraph paragraph, string search, string replaceWith)
        {
            if (paragraph.ParagraphText.Length > 0)
            {
                var parText = paragraph.ParagraphText;

                var matches = ArgumentDelimiterRegex.Matches(parText) as IReadOnlyList<Match>;

                foreach (var match in matches)
                {
                    if (!match.Success)
                    {
                        continue;
                    }

                    var matchString = match.Groups[0].Value;
                    var matchContents = match.Groups[1].Value;

                    if (search == matchContents)
                    {
                        parText = parText.Replace(matchString, replaceWith);
                    }

                }

                paragraph.ReplaceText(paragraph.ParagraphText, parText);
            }
        }

        public static int WriteToDocument(
            ref XWPFDocument document,
            string columnName,
            string columnValue,
            string nullString,
            bool append,
            out string? error)
        {
            error = null;

            try
            {
                bool isEmptyDocument = document.Paragraphs.Count == 0;
                XWPFParagraph? defaultParagraph = null;

                if (append || isEmptyDocument)
                {
                    defaultParagraph ??= document.CreateParagraph();
                    var run = defaultParagraph.CreateRun();
                    run.SetText($"{columnName}: {columnValue}");
                    run.AddBreak(BreakType.TEXTWRAPPING);
                }
                else
                {
                    foreach (var element in document.BodyElements)
                    {
                        ReplaceStringInDocument(element, columnName, string.IsNullOrEmpty(columnValue) ? nullString : columnValue);
                    }

                }
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
            return 1;
        }

        public static int WriteToDocument(
            in XWPFDocument document,
            string[] data,
            string separator,
            string nullString,
            out string? error)
        {
            error = null;

            bool isEmptyDocument = document.Paragraphs.Count == 0;
            XWPFParagraph? defaultParagraph = null;

            var arguments = ConvertToDictionary(data, separator, nullString);

            try
            {
                int i = 0;

                if (!isEmptyDocument)
                {
                    foreach (var element in document.BodyElements)
                    {
                        ProcessBodyElements(element, arguments);
                    }

                }
                else
                {
                    foreach (var target in arguments)
                    {
                        defaultParagraph ??= document.CreateParagraph();
                        var run = defaultParagraph.CreateRun();
                        run.SetText($"{target.Key}: {target.Value}");
                        run.AddBreak(BreakType.TEXTWRAPPING);

                    }
                    ++i;
                }
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;

                return -1;
            }

        }

        public static int SaveDocument(in XWPFDocument document, string path, out string? error)
        {
            try
            {
                using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                error = null;

                document.Write(stream);
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }
    }
}