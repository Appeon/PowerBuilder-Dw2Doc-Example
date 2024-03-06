namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public class DwControlMatrix
    {
        internal ISet<Dw2DObject> Controls { get; }
        internal IList<DwBand> Bands { get; }

        private DwControlMatrix()
        {
            Controls = new HashSet<Dw2DObject>();
            Bands = new List<DwBand>();
        }

        internal DwControlMatrix(
            ISet<Dw2DObject> controls, IList<DwBand> band)
        {
            Controls = controls;
            Bands = band;
        }
    }
}
