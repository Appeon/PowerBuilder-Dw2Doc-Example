using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes
{
    public class DwShapeAttributes : DwObjectAttributesBase
    {
        public Shape Shape { get; set; }
        public DwColorWrapper FillColor { get; set; } = new DwColorWrapper();
        public FillStyle FillStyle { get; set; }
        public DwColorWrapper OutlineColor { get; set; } = new DwColorWrapper();
        public LineStyle OutlineStyle { get; set; }
        public ushort OutlineWidth { get; set; } = 1;

        public DwShapeAttributes()
        {
            Floating = true;
        }
    }
}
