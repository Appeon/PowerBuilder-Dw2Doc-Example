using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes
{
    public class DwLineAttributes : DwObjectAttributesBase
    {
        public Point Start { get; set; }
        public Point End { get; set; }
        public ushort LineWidth { get; set; }
        public DwColorWrapper LineColor { get; set; } = new DwColorWrapper();
        public LineStyle LineStyle { get; set; }

        public void SetStart(int x, int y)
        {
            Start = new Point(x, y);
        }

        public void SetEnd(int x, int y)
        {
            End = new Point(x, y);
        }

        public DwLineAttributes() { }
    }
}
