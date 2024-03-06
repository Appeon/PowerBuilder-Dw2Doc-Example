using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public class VirtualCell
    {
        public Dw2DObject Object { get; protected set; }
        public RowDefinition? OwningRow { get; set; }
        public ColumnDefinition? OwningColumn { get; set; }
        public int ColumnSpan { get; set; } = 1;
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public int RightBound => X + Width;
        public int LowerBound => Y + Height;

        public VirtualCell(Dw2DObject @object)
        {
            Object = @object;
            X = Object.X;
            Y = Object.Y;
            Width = Object.Width;
            Height = Object.Height;
        }

        public override string ToString() => Object.ToString();

        public override int GetHashCode() => Object.GetHashCode();

        public override bool Equals(object? obj) => obj is VirtualCell cell && cell.Object == Object;
    }
}
