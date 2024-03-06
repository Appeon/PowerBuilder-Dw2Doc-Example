using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public class RowDefinition : EntityDefinition<RowDefinition>
    {
        public string? BandName { get; set; }
        public bool CompensatingSkew { get; set; }

        public override string? ToString()
        {
            var sb = new StringBuilder();

            sb.Append($"[Y={Offset}]");
            sb.Append($"[Height={Size}] ");
            sb.Append($"[Controls={{{Objects.Aggregate(string.Empty, (str, obj) => $"{(str == string.Empty ? str : str + ", ")} {obj}")}}}]");
            sb.Append($"[Floating={{{FloatingObjects.Aggregate(string.Empty, (str, obj) => $"{(str == string.Empty ? str : str + ", ")} {obj}")}}}]");

            return sb.ToString();
        }

        public void SortControlsByX()
        {
            Objects = Objects.OrderBy(control => control.X)
                .ToList();

            FloatingObjects = FloatingObjects.OrderBy(control => control.X)
                .ToList();
        }
    }
}
