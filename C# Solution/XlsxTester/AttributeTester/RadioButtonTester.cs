using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.Utils.CodeTable;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwRadioButtonAttributes), typeof(RadioButtonTester))]
public class RadioButtonTester : AbstractAttributeTester<DwRadioButtonAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwRadioButtonAttributes attr, ExportedCell cell)
    {

        var testResults = TestCellBase(attr, cell);

        testResults.Add(new AttributeTestResult("output cell",
            NonNullString,
            cell.OutputCell is null ? NullString : NonNullString));

        while (cell.OutputCell is not null)
        {
            testResults.Add(new AttributeTestResult("has codetable",
                bool.TrueString,
                (attr.CodeTable is var codeTable && codeTable is not null).ToString()));

            if (codeTable is null)
            {
                break;
            }

            if (attr.Text is not null)
                testResults.Add(new AttributeTestResult("text",
                    bool.TrueString,
                    (cell.OutputCell.StringCellValue == CodeTableTools.BuildString(codeTable, attr.Text, attr.LeftText, attr.Columns)).ToString()));

            XSSFCellStyle? style = cell.OutputCell.CellStyle as XSSFCellStyle;

            testResults.Add(new AttributeTestResult("cell style",
                NonNullString,
                style is null ? NullString : NonNullString));

            if (style is null)
                break;

            var font = style.GetFont();
            testResults.Add(new AttributeTestResult("font family",
                attr.FontFace,
                font.FontName));

            testResults.Add(new AttributeTestResult("font size",
                attr.FontSize.ToString(),
                (font.FontHeightInPoints - 1).ToString()));

            testResults.Add(new AttributeTestResult("font color",
                attr.FontColor.Value.ToRgb().ToString(),
                font.Color.ToString()));

            testResults.Add(new AttributeTestResult("is bold",
                attr.FontWeight.ToString(),
                font.IsBold ? "700" : "400"));

            testResults.Add(new AttributeTestResult("is italic",
                attr.Italics.ToString(),
                font.IsItalic.ToString()));

            testResults.Add(new AttributeTestResult("is strikethrough",
                attr.Strikethrough.ToString(),
                font.IsStrikeout.ToString()));

            testResults.Add(new AttributeTestResult("is underline",
                attr.Underline.ToString(),
                font.Underline.ToString()));

            testResults.Add(new AttributeTestResult("background color",
                attr.BackgroundColor.Value.ToRgb().ToString(),
                style.FillForegroundColor.ToString()));

            testResults.Add(new AttributeTestResult("background transparent",
                (attr.BackgroundColor.Value.A == 0).ToString(),
                (style.FillPattern == NPOI.SS.UserModel.FillPattern.NoFill).ToString()));

            testResults.Add(new AttributeTestResult("horizontal alignment",
                attr.Alignment.ToNpoiHorizontalAlignment().ToString(),
                style.Alignment.ToString()));

            testResults.Add(new AttributeTestResult("left text",
                attr.LeftText.ToString(),
                (cell.OutputCell.StringCellValue[0] switch
                {
                    '✅' => true,
                    '☐' => true,
                    _ => false
                }).ToString()));

            break;
        }

        return new AttributeTestResultCollection(testResults);
    }

    protected override AttributeTestResultCollection TestFloating(DwRadioButtonAttributes attr, ExportedFloatingCell cell)
    {
        var testResults = TestFloatingBase(attr, cell);

        testResults.Add(new("output shape", NonNullString, cell.OutputShape is null ? NullString : NonNullString));

        while (cell.OutputShape is not null)
        {
            testResults.Add(new AttributeTestResult("has codetable",
                bool.TrueString,
                (attr.CodeTable is var codeTable && codeTable is not null).ToString()));

            if (codeTable is null)
            {
                break;
            }

            testResults.Add(new("shape type", nameof(XSSFTextBox), cell.OutputShape.GetType().Name));
            if (cell.OutputShape is not XSSFTextBox textBox)
            {
                break;
            }

            if (attr.BackgroundColor.Value.A != 0)
            {
                testResults.Add(new("background color",
                    attr.BackgroundColor.Value.ToArgb().ToString(),
                    textBox.FillColor.ToString()));
            }

            testResults.Add(new("paragraphs",
                 "1",
                 textBox.TextParagraphs.Count.ToString()));

            if (textBox.TextParagraphs.Count == 0)
                break;

            var paragraph = textBox.TextParagraphs[0];

            testResults.Add(new("text run",
                "1",
                paragraph.TextRuns.Count.ToString()));

            if (paragraph.TextRuns.Count == 0) break;


            var run = paragraph.TextRuns[0];

            testResults.Add(new("text run text",
                CodeTableTools.BuildString(codeTable, attr.Text, attr.LeftText, attr.Columns),
                run.Text));

            testResults.Add(new("font color",
                attr.FontColor.Value.ToRgb().ToString(),
                (run.FontColor.R << 16 + run.FontColor.G << 8 + run.FontColor.B).ToString()));


            testResults.Add(new AttributeTestResult("left text",
                attr.LeftText.ToString(),
                (run.Text[0] switch
                {
                    '◉' or '⊙' => true,
                    _ => false
                }).ToString()));

            testResults.Add(new AttributeTestResult("font size",
                attr.FontSize.ToString(),
                (run.FontSize - 1).ToString()));

            testResults.Add(new AttributeTestResult("is bold",
                attr.FontWeight.ToString(),
                run.IsBold ? "700" : "400"));

            testResults.Add(new AttributeTestResult("is italic",
                attr.Italics.ToString(),
                run.IsItalic.ToString()));

            testResults.Add(new AttributeTestResult("is strikethrough",
                attr.Strikethrough.ToString(),
                run.IsStrikethrough.ToString()));

            testResults.Add(new AttributeTestResult("is underline",
                attr.Underline.ToString(),
                run.IsUnderline.ToString()));

            testResults.Add(new AttributeTestResult("alignment",
                attr.Alignment.ToNpoiTextAlignment().ToString(),
                paragraph.TextAlign.ToString()));


            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
