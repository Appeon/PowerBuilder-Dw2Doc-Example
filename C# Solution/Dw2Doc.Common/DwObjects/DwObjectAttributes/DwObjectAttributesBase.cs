namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes
{
    public abstract class DwObjectAttributesBase
    {
        public bool IsVisible { get; set; }
        public bool Floating { get; protected set; }

        public DwObjectAttributesBase()
        {
        }
    }
}
