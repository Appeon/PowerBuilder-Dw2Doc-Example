using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public class DwBand
    {
        public string Name { get; }
        public int Height { get; init; }
        public int Position { get; set; }
        public int Bound => Height + Position;
        public bool Repeats { get; }

        public BandType BandType { get; set; }

        public DwBand? ParentBand { get; set; }
        public DwBand? RelatedHeader { get; set; }
        public IList<DwBand> RelatedTrailers { get; }

        public DwBand(string name, bool repeats)
        {
            RelatedTrailers = new List<DwBand>();
            Name = name;
            Repeats = repeats;
        }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.Append($"{nameof(DwBand)} [Name={Name}][Height={Height}][Position={Position}]");

            return sb.ToString();
        }

        public override bool Equals(object? obj)
        {
            return obj is DwBand b && b.Name == Name;
        }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }
    }
}
