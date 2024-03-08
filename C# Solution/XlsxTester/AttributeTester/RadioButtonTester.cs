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

        testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "output cell",
            NonNullString,
            cell.OutputCell is null ? NullString : NonNullString));

        while (cell.OutputCell is not null)
        {
            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "has codetable",
                bool.TrueString,
                (attr.CodeTable is var codeTable && codeTable is not null).ToString()));

            if (codeTable is null)
            {
                break;
            }

            if (attr.Text is not null)
                testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "text",
                    bool.TrueString,
                    (cell.OutputCell.StringCellValue == CodeTableTools.BuildString(codeTable, attr.Text, attr.LeftText, attr.Columns)).ToString()));

            XSSFCellStyle? style = cell.OutputCell.CellStyle as XSSFCellStyle;

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "cell style",
                NonNullString,
                style is null ? NullString : NonNullString));

            if (style is null)
                break;

            var font = style.GetFont();
            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "font family",
                attr.FontFace,
                font.FontName));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "font size",
                attr.FontSize.ToString(),
                (font.FontHeightInPoints + 1).ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "font color",
                attr.FontColor.Value.ToRgb().ToString(),
                font.Color.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is bold",
                attr.FontWeight.ToString(),
                font.IsBold ? "700" : "400"));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is italic",
                attr.Italics.ToString(),
                font.IsItalic.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is strikethrough",
                attr.Strikethrough.ToString(),
                font.IsStrikeout.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is underline",
                attr.Underline.ToString(),
                font.Underline.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "background color",
                attr.BackgroundColor.Value.ToRgb().ToString(),
                style.FillForegroundColor.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "background transparent",
                (attr.BackgroundColor.Value.A == 0).ToString(),
                (style.FillPattern == NPOI.SS.UserModel.FillPattern.NoFill).ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "horizontal alignment",
                attr.Alignment.ToNpoiHorizontalAlignment().ToString(),
                style.Alignment.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "left text",
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

        testResults.Add(new(
            cell.Cell.Object.Name,
            "output shape",
            NonNullString,
            cell.OutputShape is null ? NullString : NonNullString));

        while (cell.OutputShape is not null)
        {
            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "has codetable",
                bool.TrueString,
                (attr.CodeTable is var codeTable && codeTable is not null).ToString()));

            if (codeTable is null)
            {
                break;
            }

            testResults.Add(new(
                cell.Cell.Object.Name,
                "shape type",
                nameof(XSSFTextBox),
                cell.OutputShape.GetType().Name));
            if (cell.OutputShape is not XSSFTextBox textBox)
            {
                break;
            }

            if (attr.BackgroundColor.Value.A != 0)
            {
                testResults.Add(new(
                    cell.Cell.Object.Name,
                    "background color",
                    attr.BackgroundColor.Value.ToArgb().ToString(),
                    textBox.FillColor.ToString()));
            }

            testResults.Add(new(
                cell.Cell.Object.Name,
                "paragraphs",
                 "1",
                 textBox.TextParagraphs.Count.ToString()));

            if (textBox.TextParagraphs.Count == 0)
                break;

            var paragraph = textBox.TextParagraphs[0];

            testResults.Add(new(
                cell.Cell.Object.Name,
                "text run",
                "1",
                paragraph.TextRuns.Count.ToString()));

            if (paragraph.TextRuns.Count == 0) break;


            var run = paragraph.TextRuns[0];

            testResults.Add(new(
                cell.Cell.Object.Name,
                "text run text",
                CodeTableTools.BuildString(codeTable, attr.Text, attr.LeftText, attr.Columns),
                run.Text));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font color",
                attr.FontColor.Value.ToRgb().ToString(),
                (run.FontColor.R << 16 + run.FontColor.G << 8 + run.FontColor.B).ToString()));


            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "left text",
                attr.LeftText.ToString(),
                (run.Text[0] switch
                {
                    '◉' or '⊙' => false,
                    _ => true
                }).ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "font size",
                attr.FontSize.ToString(),
                (run.FontSize + 1).ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is bold",
                attr.FontWeight.ToString(),
                run.IsBold ? "700" : "400"));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is italic",
                attr.Italics.ToString(),
                run.IsItalic.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is strikethrough",
                attr.Strikethrough.ToString(),
                run.IsStrikethrough.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "is underline",
                attr.Underline.ToString(),
                run.IsUnderline.ToString()));

            testResults.Add(new AttributeTestResult(cell.Cell.Object.Name, "alignment",
                attr.Alignment.ToNpoiTextAlignment().ToString(),
                paragraph.TextAlign.ToString()));


            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
