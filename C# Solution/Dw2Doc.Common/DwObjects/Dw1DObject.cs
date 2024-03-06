namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    internal class Dw1DObject : Dw2DObject
    {
        public int X1 { get; }
        public int X2 { get; }
        public int Y1 { get; }
        public int Y2 { get; }

        public Dw1DObject(string name, string band, int x1, int x2, int y1, int y2)
            : base(name,
                  band,
                  Math.Min(x1, x2),
                  Math.Min(y1, y2),
                  Math.Abs(x2 - x1),
                  Math.Abs(y2 - y1))
        {
            X1 = x1;
            X2 = x2;
            Y1 = y1;
            Y2 = y2;
        }
    }
}
