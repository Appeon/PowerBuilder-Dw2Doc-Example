using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

public abstract class AbstractAttributeTester<TAttr>
    where TAttr : DwObjectAttributesBase
{
    protected const string NullString = "null";
    protected const string NonNullString = "non-null value";

    public AttributeTestResultCollection Test(TAttr attr, ExportedCellBase exportedCell) => exportedCell switch
    {
        ExportedCell cell => TestCell(attr, cell),
        ExportedFloatingCell shape => TestFloating(attr, shape),
        _ => throw new NotImplementedException()
    };

    protected abstract AttributeTestResultCollection TestCell(TAttr attr, ExportedCell cell);

    protected abstract AttributeTestResultCollection TestFloating(TAttr attr, ExportedFloatingCell cell);

    protected List<AttributeTestResult> TestCellBase(TAttr attr, ExportedCell cell)
    {
        List<AttributeTestResult> results = new()
        {
            //new AttributeTestResult("floating", bool.FalseString, bool.FalseString),
        };

        return results;
    }

    protected List<AttributeTestResult> TestFloatingBase(TAttr attr, ExportedFloatingCell cell)
    {
        List<AttributeTestResult> results = new()
        {
            //new AttributeTestResult("floating", bool.TrueString, attr.Floating.ToString().ToLower())
        };

        return results;
    }

}
