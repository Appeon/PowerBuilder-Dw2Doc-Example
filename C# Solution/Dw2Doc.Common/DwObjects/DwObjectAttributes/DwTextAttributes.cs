using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes
{
    public class DwTextAttributes : DwObjectAttributesBase
    {
        public string? Text { get; set; }
        public Alignment Alignment { get; set; } = Alignment.Left;
        public byte FontSize { get; set; }
        public short FontWeight { get; set; }
        public bool Underline { get; set; }
        public bool Italics { get; set; }
        public bool Strikethrough { get; set; }
        public string? FontFace { get; set; }
        public DwColorWrapper FontColor { get; set; } = new DwColorWrapper()
        {
            Value = Color.FromArgb(255, 0, 0, 0)
        };
        public DwColorWrapper BackgroundColor { get; set; } = new DwColorWrapper()
        {
            Value = Color.FromArgb(0, 255, 255, 255)
        };

        public DwTextAttributes()
        {
            Floating = false;
        }

        public override bool Equals(object? obj)
        {
            return obj is DwTextAttributes other && Text == other.Text;
        }

        public override int GetHashCode()
        {
            return Text?.GetHashCode() ?? 0;
        }

        public override string ToString()
        {
            return $"{Text}";
        }
    }
}
