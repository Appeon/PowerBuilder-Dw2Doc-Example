using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public class FloatingVirtualCell : VirtualCell
    {
        public FloatingCellOffset Offset { get; }

        public FloatingVirtualCell(Dw2DObject @object, FloatingCellOffset offset) : base(@object)
        {
            Offset = offset;
        }

        public static FloatingVirtualCell FromVirtualCell(VirtualCell cell, FloatingCellOffset offset)
            => new(cell.Object, offset)
            {
                OwningColumn = cell.OwningColumn,
                OwningRow = offset.StartRow,
            };
    }
}
