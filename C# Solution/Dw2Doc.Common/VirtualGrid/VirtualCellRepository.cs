using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    internal class VirtualCellRepository
    {
        internal Dictionary<string, VirtualCell> CellsByControlName { get; }
        internal Dictionary<string, IDictionary<string, VirtualCell>> RowsByBand { get; }
        internal SortedList<int, ISet<VirtualCell>> CellsByX { get; }
        internal SortedList<int, IList<VirtualCell>> CellsByY { get; }
        internal HashSet<VirtualCell> Cells { get; }

        private VirtualCellRepository()
        {
            CellsByControlName = new Dictionary<string, VirtualCell>();
            RowsByBand = new Dictionary<string, IDictionary<string, VirtualCell>>();
            CellsByX = new SortedList<int, ISet<VirtualCell>>();
            CellsByY = new SortedList<int, IList<VirtualCell>>();
            Cells = new HashSet<VirtualCell>();
        }

        internal VirtualCellRepository(Dictionary<string, VirtualCell> controlsByName,
            Dictionary<string, IDictionary<string, VirtualCell>> cellsByBand,
            SortedList<int, ISet<VirtualCell>> controlsByX,
            SortedList<int, IList<VirtualCell>> controlsByY,
            HashSet<VirtualCell> controls)
        {
            CellsByControlName = controlsByName;
            RowsByBand = cellsByBand;
            CellsByX = controlsByX;
            CellsByY = controlsByY;
            Cells = controls;
        }

        public override string ToString()
        {
            StringBuilder sb = new();
            sb.Append("X order]: ");
            foreach (var set in CellsByX)
            {
                foreach (var control in set.Value)
                {
                    sb.Append($"{control.Object.Name}({control.X})");
                    sb.Append(',');
                }
                sb.Length--;
                sb.Append(':');
            }
            sb.Length--;

            sb.AppendLine();
            sb.Append("Y order]: ");
            foreach (var set in CellsByY)
            {
                foreach (var control in set.Value)
                {
                    sb.Append($"{control.Object.Name} ( {control.Y})");
                    sb.Append(',');
                }
                sb.Length--;
                sb.Append(':');
            }
            sb.Length--;


            return sb.ToString();
        }
    }
}
