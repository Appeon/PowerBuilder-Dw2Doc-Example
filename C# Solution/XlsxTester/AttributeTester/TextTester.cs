using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwTextAttributes), typeof(TextTester))]
public class TextTester : AbstractAttributeTester<DwTextAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwTextAttributes attr, ExportedCell cell)
    {
        var testResults = TestCellBase(attr, cell);

        testResults.Add(new(
            cell.Cell.Object.Name,
                "output cell",
                NonNullString,
                cell.OutputCell is null ? NullString : NonNullString
            ));

        while (cell.OutputCell is not null)
        {
            var cellStyle = cell.OutputCell.CellStyle as XSSFCellStyle;
            testResults.Add(new(
                cell.Cell.Object.Name,
                "has style",
                bool.TrueString,
                (cellStyle is not null).ToString()
            ));

            if (cellStyle is null)
            {
                break;
            }

            var font = cellStyle.GetFont();
            testResults.Add(new(
                cell.Cell.Object.Name,
                "has font",
                bool.TrueString,
                (font is not null).ToString()
            ));

            if (font is null)
                break;

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font height",
                (attr.FontSize - 1).ToString(),
                font.FontHeightInPoints.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font name",
                attr.FontFace,
                font.FontName
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font color",
                attr.FontColor.Value.ToRgb().ToString(),
                font.Color.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is bold",
                (attr.FontWeight >= 700).ToString(),
                font.IsBold.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is italic",
                attr.Italics.ToString(),
                font.IsItalic.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is strikethrough",
                attr.Strikethrough.ToString(),
                font.IsStrikeout.ToString()
            ));

            bool transparentBg = (attr.BackgroundColor.Value.A == 0);

            testResults.Add(new(
                cell.Cell.Object.Name,
                "background transparent",
                transparentBg.ToString(),
                (cellStyle.FillPattern == NPOI.SS.UserModel.FillPattern.NoFill).ToString()
            ));

            if (!transparentBg)
                testResults.Add(new(
                    cell.Cell.Object.Name,
                    "background color",
                    attr.BackgroundColor.Value.ToRgb().ToString(),
                    cellStyle.FillForegroundColor.ToString()
                ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "alignment",
                attr.Alignment.ToNpoiHorizontalAlignment().ToString().ToLower(),
                cellStyle.Alignment.ToString().ToLower()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "text",
                attr.Text,
                cell.OutputCell.StringCellValue
            ));

            break;
        }

        return new AttributeTestResultCollection(testResults);
    }

    protected override AttributeTestResultCollection TestFloating(DwTextAttributes attr, ExportedFloatingCell cell)
    {
        var testResults = TestFloatingBase(attr, cell);

        testResults.Add(new(
            cell.Cell.Object.Name,
                "output shape",
                NonNullString,
                cell.OutputShape is null ? NullString : NonNullString
            ));

        while (cell.OutputShape is not null)
        {
            testResults.Add(new(
                cell.Cell.Object.Name,
                "shape type",
                nameof(XSSFTextBox),
                cell.OutputShape.GetType().Name
            ));

            if (cell.OutputShape is not XSSFTextBox textBox) { break; }

            /// Don't test FillColor because its getter is not implemented in NPOI
            //testResults.Add(new(
            //    "background color",
            //    attr.BackgroundColor.Value.ToRgb().ToString(),
            //    textBox.FillColor.ToString()
            //));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "paragraphs",
                "1",
                textBox.TextParagraphs.Count.ToString()
            ));

            if (textBox.TextParagraphs.Count == 0)
                break;

            var paragraph = textBox.TextParagraphs[0];

            testResults.Add(new(
                cell.Cell.Object.Name,
                "text run",
                "1",
                paragraph.TextRuns.Count.ToString()
            ));

            if (paragraph.TextRuns.Count == 0)
                break;

            var run = paragraph.TextRuns[0];

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font color",
                attr.FontColor.Value.ToRgb().ToString(),
                (run.FontColor.R << 16 + run.FontColor.G << 8 + run.FontColor.B).ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "text",
                attr.Text,
                run.Text
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font size",
                (attr.FontSize - 1).ToString(),
                run.FontSize.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is italic",
                attr.Italics.ToString(),
                run.IsItalic.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is strikethrough",
                attr.Strikethrough.ToString(),
                run.IsStrikethrough.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is underline",
                attr.Underline.ToString(),
                run.IsUnderline.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "is bold",
                (attr.FontWeight >= 700).ToString(),
                run.IsBold.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "font name",
                attr.FontFace,
                run.FontFamily
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "alignment",
                attr.Alignment.ToNpoiHorizontalAlignment().ToString().ToLower(),
                paragraph.TextAlign.ToString().ToLower()
            ));

            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
