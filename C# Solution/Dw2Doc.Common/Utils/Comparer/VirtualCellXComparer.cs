using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Utils.Comparer
{
    public class VirtualCellXComparer : IComparer<VirtualCell>
    {
        public int Compare(VirtualCell? x, VirtualCell? y)
        {
            if (x == null)
                throw new ArgumentNullException(nameof(x));
            if (y == null)
                throw new ArgumentNullException(nameof(y));
            return x.X.CompareTo(y.X);
        }
    }
}
