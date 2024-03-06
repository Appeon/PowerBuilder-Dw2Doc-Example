using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;
using Appeon.DotnetDemo.Dw2Doc.Common.Utils.Comparer;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Extensions
{
    public static class CollectionExtensions
    {
        public static void AddRange<T>(this ICollection<T> destination,
            IEnumerable<T> source)
        {
            foreach (var item in source)
            {
                destination.Add(item);
            }
        }

        public static string CollectionToString<T>(this ICollection<T> source, string separator)
        {
            StringBuilder sb = new();
            bool flag = false;
            foreach (var item in source)
            {
                flag = true;
                if (item is not null && item.GetType() == typeof(ICollection<>))
                {
                    sb.AppendLine(CollectionToString(source, separator));
                }
                else
                    sb.AppendLine(item?.ToString());
            }
            if (flag)
                sb.Length -= separator.Length;

            return sb.ToString();
        }

        public static string CollectionToString<T>(this ICollection<T> source)
        {
            return CollectionToString(source, ", ");
        }

        internal static VirtualCellRepository ToVirtualCellRepository(this IEnumerable<Dw2DObject> controls)
        {
            var byName = new Dictionary<string, VirtualCell>();
            var byBand = new Dictionary<string, IDictionary<string, VirtualCell>>();
            var byY = new SortedList<int, IList<VirtualCell>>();
            var byX = new SortedList<int, ISet<VirtualCell>>();
            var cells = new HashSet<VirtualCell>();

            VirtualCell cell;
            foreach (var control in controls)
            {
                cell = new VirtualCell(control);
                byName[control.Name] = cell;
                if (!byX.ContainsKey(control.X))
                    byX[control.X] = new HashSet<VirtualCell>();

                byX[control.X].Add(cell);

                if (!byY.ContainsKey(control.Y))
                    byY[control.Y] = new List<VirtualCell>();

                if (!byBand.ContainsKey(control.Band))
                    byBand[control.Band] = new Dictionary<string, VirtualCell>();

                byBand[control.Band][control.Name] = cell;

                byY[control.Y].Add(cell);
                cells.Add(cell);
            }

            foreach (var (_, controlSet) in byY)
            {
                (controlSet as List<VirtualCell>)!.Sort(new VirtualCellXComparer());
            }

            return new VirtualCellRepository(
                byName,
                byBand,
                byX,
                byY,
                cells);
        }

        internal static IList<BandRows> DwBandsToBandRows(this IEnumerable<DwBand> bands, IDictionary<string, IList<RowDefinition>> rowsPerBand)
        {
            var bandRows = new List<BandRows>();

            var bandRowNameMap = new Dictionary<string, BandRows>();
            BandRows newBandRow;
            foreach (var row in bands)
            {
                newBandRow = rowsPerBand.ContainsKey(row.Name)
                    ? new BandRows(row.Name, rowsPerBand[row.Name])
                    : new BandRows(row.Name, new List<RowDefinition>());
                newBandRow.IsRepeatable = row.Repeats;
                newBandRow.BandType = row.BandType;
                bandRows.Add(newBandRow);
                bandRowNameMap[row.Name] = newBandRow;
            }

            foreach (var band in bands)
            {
                foreach (var relatedTrailer in band.RelatedTrailers)
                {
                    bandRowNameMap[band.Name].RelatedTrailers.Add(bandRowNameMap[relatedTrailer.Name]);
                }

                if (band.ParentBand is not null)
                    bandRowNameMap[band.Name].ParentBand = bandRowNameMap[band.ParentBand.Name];
                if (band.RelatedHeader is not null)
                    bandRowNameMap[band.Name].RelatedHeader = bandRowNameMap[band.RelatedHeader.Name];
            }


            return bandRows;
        }

        public static void AssignToEntity<T>(this IEnumerable<VirtualCell> cells,
            EntityDefinition<T> targetEntity)
            where T : EntityDefinition<T>
        {
            if (targetEntity is RowDefinition row)
            {
                foreach (var cell in cells)
                {
                    cell.OwningRow = row;
                }
            }
            else if (targetEntity is ColumnDefinition col)
            {
                foreach (var cell in cells)
                {
                    cell.OwningColumn = col;
                }
            }
        }
    }
}
