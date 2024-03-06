using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjectAttributes
{
    public class DwAttributeFactory
    {
        public static DwTextAttributes CreateTextAttributes() => new();
        public static DwTextAttributes CreateTextAttributes(
            bool isVisible,
            int alignment,
            int color,
            int bgColor,
            string fontFace,
            byte fontSize,
            short fontWeight,
            bool italics,
            bool underline,
            bool strikethrough,
            string text) => new()
            {
                Alignment = (Enums.Alignment)alignment,
                FontColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        color & 0xFF,
                        (color & 0xFF00) >> 8,
                        (color & 0xFF0000) >> 16
                        )
                },
                BackgroundColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        (bgColor & 0x20000000) > 0 ? 0 : 255,
                        bgColor & 0xFF,
                        (bgColor & 0xFF00) >> 8,
                        (bgColor & 0xFF0000) >> 16
                        )
                },
                FontFace = fontFace,
                FontSize = fontSize,
                FontWeight = fontWeight,
                Italics = italics,
                Underline = underline,
                Strikethrough = strikethrough,
                IsVisible = isVisible,
                Text = text
            };

        public static DwCheckboxAttributes CreateCheckboxAttribute() => new();
        public static DwCheckboxAttributes CreateCheckboxAttribute(
            bool isVisible,
            int alignment,
            int color,
            int bgColor,
            string fontFace,
            byte fontSize,
            short fontWeight,
            bool italics,
            bool underline,
            bool strikethrough,
            string text,
            string label,
            string checkedValue,
            string uncheckedValue) => new()
            {
                Alignment = (Enums.Alignment)alignment,
                FontColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        color & 0xFF,
                        (color & 0xFF00) >> 8,
                        (color & 0xFF0000) >> 16
                        )
                },
                BackgroundColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        (bgColor & 0x20000000) > 0 ? 0 : 255,
                        bgColor & 0xFF,
                        (bgColor & 0xFF00) >> 8,
                        (bgColor & 0xFF0000) >> 16
                        )
                },
                FontFace = fontFace,
                FontSize = fontSize,
                FontWeight = fontWeight,
                Italics = italics,
                Underline = underline,
                Strikethrough = strikethrough,
                IsVisible = isVisible,
                Text = text,
                Label = label,
                CheckedValue = checkedValue,
                UncheckedValue = uncheckedValue,
            };

        public static DwRadioButtonAttributes CreateRadioButtonAttribute() => new();
        public static DwRadioButtonAttributes CreateRadioButtonAttribute(
            bool isVisible,
            int alignment,
            int color,
            int bgColor,
            string fontFace,
            byte fontSize,
            short fontWeight,
            bool italics,
            bool underline,
            bool strikethrough,
            string text,
            IDictionary<string, string> codemap,
            short columns) => new()
            {
                Alignment = (Enums.Alignment)alignment,
                FontColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        color & 0xFF,
                        (color & 0xFF00) >> 8,
                        (color & 0xFF0000) >> 16
                        )
                },
                BackgroundColor = new DwObjects.DwColorWrapper()
                {
                    Value = Color.FromArgb(
                        (bgColor & 0x20000000) > 0 ? 0 : 255,
                        bgColor & 0xFF,
                        (bgColor & 0xFF00) >> 8,
                        (bgColor & 0xFF0000) >> 16
                        )
                },
                FontFace = fontFace,
                FontSize = fontSize,
                FontWeight = fontWeight,
                Italics = italics,
                Underline = underline,
                Strikethrough = strikethrough,
                IsVisible = isVisible,
                Text = text,
                CodeTable = codemap,
                Columns = columns
            };



        public static DwShapeAttributes CreateShapeAttributes() => new();
        public static DwShapeAttributes CreateShapeAttributes(
            bool isVisible,
            int fillColor,
            byte fillStyle,
            int outlineColor,
            byte outlineStyle,
            ushort outlineWidth,
            byte shape) => new()
            {
                FillColor = new DwObjects.DwColorWrapper() { Value = Color.FromArgb((fillColor & 0xFF0000) >> 16, (fillColor & 0xFF00) >> 8, fillColor & 0xFF) },
                FillStyle = (Enums.FillStyle)fillStyle,
                IsVisible = isVisible,
                OutlineColor = new DwObjects.DwColorWrapper() { Value = Color.FromArgb((outlineColor & 0xFF0000) >> 16, (outlineColor & 0xFF00) >> 8, outlineColor & 0xFF) },
                OutlineStyle = (Enums.LineStyle)outlineStyle,
                OutlineWidth = outlineWidth,
                Shape = (Enums.Shape)shape,
            };


        public static DwLineAttributes CreateLineAttributes() => new();
        public static DwLineAttributes CreateLineAttributes(
            bool isVisible,
            int startPointX,
            int startPointY,
            int endPointX,
            int endPointY,
            int lineColor,
            byte lineStyle,
            ushort lineWidth) => new()
            {
                IsVisible = isVisible,
                LineColor = new DwObjects.DwColorWrapper() { Value = Color.FromArgb((lineColor & 0xFF0000) >> 16, (lineColor & 0xFF00) >> 8, lineColor & 0xFF) },
                LineStyle = (Enums.LineStyle)lineStyle,
                LineWidth = lineWidth,
                Start = new Point(startPointX, startPointY),
                End = new Point(endPointX, endPointY),
            };


        public static DwPictureAttributes CreatePictureAttributes() => new();
        public static DwPictureAttributes CreatePictureAttributes(
            bool isVisible,
            string fileName,
            byte transparency) => new()
            {
                FileName = fileName,
                IsVisible = isVisible,
                Transparency = transparency
            };

        public static DwButtonAttributes CreateButtonAttributes() => new();
        public static DwButtonAttributes CreateButtonAttributes(
            bool isVisible,
            int fontSize,
            string text) => new()
            {
                FontSize = fontSize,
                IsVisible = isVisible,
                Text = text
            };
    }
}

