using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwButtonAttributes), typeof(ButtonAttributeTester))]
public class ButtonAttributeTester : AbstractAttributeTester<DwButtonAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwButtonAttributes attr, ExportedCell cell)
    {
        throw new InvalidOperationException("Button objects should not be cells");
    }

    protected override AttributeTestResultCollection TestFloating(DwButtonAttributes attr, ExportedFloatingCell cell)
    {

        var attributeTests = TestFloatingBase(attr, cell);


        attributeTests.Add(new AttributeTestResult(cell.Cell.Object.Name,
            "output shape",
            NonNullString,
            cell.OutputShape is null ? NullString : NonNullString));

        while (cell.OutputShape is not null)
        {

            attributeTests.Add(new AttributeTestResult(cell.Cell.Object.Name,
                "shapetype",
                nameof(XSSFSimpleShape),
                cell.OutputShape.GetType().Name));

            var simpleShape = (cell.OutputShape as XSSFSimpleShape)!;

            attributeTests.Add(new AttributeTestResult(cell.Cell.Object.Name,
                "text paragraphs",
                "1",
                simpleShape.TextParagraphs.Count.ToString()));

            if (simpleShape.TextParagraphs.Count == 0)
                break;

            var text = simpleShape.TextParagraphs[0];

            attributeTests.Add(new AttributeTestResult(cell.Cell.Object.Name,
                "text runs",
                "1",
                text.TextRuns.Count.ToString()));

            if (text.TextRuns.Count == 0)
                break;

            var textRun = text.TextRuns[0];

            attributeTests.Add(new AttributeTestResult(cell.Cell.Object.Name,
                "text",
                attr.Text,
                textRun.Text));

            break;
        }

        return new AttributeTestResultCollection(attributeTests);
    }
}
