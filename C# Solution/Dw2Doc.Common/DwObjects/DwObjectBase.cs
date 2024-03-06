namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public abstract class DwObjectBase
    {
        public string Name { get; }
        public string Band { get; }
        public bool Floating { get; set; }

        public DwObjectBase(string name, string band)
        {
            Name = name;
            Band = band;
        }

        public override bool Equals(object? obj)
        {
            return obj is DwObjectBase other && Name == other.Name;
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }

        public override string ToString()
        {
            return $"DwObjectBase[{Name}]";
        }
    }
}
