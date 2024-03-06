using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwLineAttributes), typeof(LineTester))]
public class LineTester : AbstractAttributeTester<DwLineAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwLineAttributes attr, ExportedCell cell)
    {
        throw new NotImplementedException();
    }

    protected override AttributeTestResultCollection TestFloating(DwLineAttributes attr, ExportedFloatingCell cell)
    {
        var testResults = TestFloatingBase(attr, cell);

        testResults.Add(new("output shape", NonNullString, cell.OutputShape is null ? NullString : NonNullString));

        while (cell.OutputShape is not null)
        {
            testResults.Add(new("shape type", nameof(XSSFShape), cell.OutputShape.GetType().Name));
            testResults.Add(new("line style",
                attr.LineStyle.DwLineStyleToNpoiLineStyle().ToString(),
                cell.OutputShape.LineStyle.ToString()));

            testResults.Add(new("line width",
                attr.LineWidth.ToString(),
                cell.OutputShape.LineWidth.ToString()));

            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
