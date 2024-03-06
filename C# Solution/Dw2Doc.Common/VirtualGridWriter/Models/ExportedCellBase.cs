using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;

public abstract class ExportedCellBase
{
    public VirtualCell Cell { get; init; }
    public DwObjectAttributesBase Attribute { get; init; }
    public Type AttributeType => Attribute.GetType();

    protected ExportedCellBase(
        VirtualCell cell,
        DwObjectAttributesBase attribute
        )
    {
        Cell = cell;
        Attribute = attribute;
    }
}
