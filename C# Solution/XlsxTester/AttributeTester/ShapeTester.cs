using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Models;
using Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

[TesterForAttribute(typeof(DwShapeAttributes), typeof(ShapeTester))]
public class ShapeTester : AbstractAttributeTester<DwShapeAttributes>
{
    protected override AttributeTestResultCollection TestCell(DwShapeAttributes attr, ExportedCell cell)
    {
        throw new NotImplementedException();
    }

    protected override AttributeTestResultCollection TestFloating(DwShapeAttributes attr, ExportedFloatingCell cell)
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
                nameof(XSSFSimpleShape),
                cell.OutputShape.GetType().Name
            ));

            if (cell.OutputShape is not XSSFSimpleShape shape)
            {
                break;
            }

            testResults.Add(new(
                cell.Cell.Object.Name,
                "fill color",
                attr.FillColor.Value.ToRgb().ToString(),
                shape.FillColor.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "fill transparent",
                (attr.FillStyle == Common.Enums.FillStyle.Transparent).ToString(),
                shape.IsNoFill.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "line style",
                attr.OutlineStyle.DwLineStyleToNpoiLineStyle().ToString(),
                shape.LineStyle.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "line color",
                attr.OutlineColor.Value.ToRgb().ToString(),
                shape.LineStyleColor.ToString()
            ));

            testResults.Add(new(
                cell.Cell.Object.Name,
                "line width",
                attr.OutlineWidth.ToString(),
                shape.LineWidth.ToString()
            ));


            testResults.Add(new(
                cell.Cell.Object.Name,
                "shape type",
                ((int)(attr.Shape switch
                {
                    Common.Enums.Shape.Circle => ShapeTypes.Ellipse,
                    Common.Enums.Shape.Rectangle => ShapeTypes.Rectangle,
                    _ => throw new NotImplementedException("Unsupported shape type")
                })).ToString(),
                shape.ShapeType.ToString()
            ));


            break;
        }

        return new AttributeTestResultCollection(testResults);
    }
}
