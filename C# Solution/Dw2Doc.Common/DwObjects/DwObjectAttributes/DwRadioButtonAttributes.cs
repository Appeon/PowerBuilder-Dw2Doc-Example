namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;

public class DwRadioButtonAttributes : DwTextAttributes
{
    public IDictionary<string, string>? CodeTable { get; set; }
    public short Columns { get; set; }
    public bool LeftText { get; set; }
}
