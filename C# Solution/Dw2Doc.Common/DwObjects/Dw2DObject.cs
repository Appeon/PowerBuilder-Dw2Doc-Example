namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public class Dw2DObject : DwObjectBase
    {
        public int X { get; }
        public int Y { get; set; }
        public int Width { get; }
        public int Height { get; }

        public int RightBound => X + Width;
        public int LowerBound => Y + Height;

        public Dw2DObject(string name, string band, int x, int y, int width, int height) : base(name, band)
        {
            X = x;
            Y = y;
            Height = height;
            Width = width;
        }

    }
}
