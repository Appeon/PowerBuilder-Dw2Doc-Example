using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.Extensions;
using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public class VirtualGrid
    {
        public IList<RowDefinition> Rows { get; }
        public IList<ColumnDefinition> Columns { get; }
        public IList<BandRows> BandRows { get; }
        internal VirtualCellRepository CellRepository { get; }
        public DwType DwType { get; set; }

        internal VirtualGrid(
            IList<RowDefinition> rows,
            IList<ColumnDefinition> columns,
            IList<BandRows> rowsPerBand,
            VirtualCellRepository cellRepository,
            DwType DwType
            )
        {
            Rows = rows;
            Columns = columns;
            BandRows = rowsPerBand;
            CellRepository = cellRepository;
            this.DwType = DwType;
        }

        public override string ToString()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Rows ->");
            builder.AppendLine(Rows?.CollectionToString() ?? "null");

            builder.AppendLine("Columns -> ");
            builder.AppendLine(Columns?.CollectionToString() ?? "null");

            builder.AppendLine("Rows/Band-> ");
            foreach (var bandRow in BandRows)
            {
                builder.AppendLine($"--- {bandRow.Name} ---");
                builder.AppendLine(bandRow.Rows.CollectionToString());
            }

            return builder.ToString();
        }
    }
}
