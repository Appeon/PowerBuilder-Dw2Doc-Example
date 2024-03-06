using System.Diagnostics.SymbolStore;

namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public class DwControlMatrixBuilder
    {
        private readonly ICollection<DwObjectBase> _objects;
        private IList<DwBand>? _bands;

        public DwControlMatrixBuilder()
        {
            _objects = new HashSet<DwObjectBase>();
        }

        public DwControlMatrixBuilder Add2DControl(string name, string band, int x, int y, int width, int height, bool floating = false)
        {
            var control = new Dw2DObject(name, band, x, y, width, height)
            {
                Floating = floating,
            };

            _objects.Add(control);

            return this;
        }

        public DwControlMatrixBuilder Add1DControl(string name, string band, int x1, int y1, int x2, int y2, bool floating = false)
        {
            var control = new Dw1DObject(name, band, x1, x2, y1, y2)
            {
                Floating = floating,
            };

            _objects.Add(control);

            return this;
        }

        public DwControlMatrixBuilder SetBands(IList<DwBand> bands)
        {
            _bands = bands;
            return this;
        }


        public DwControlMatrix? Build(out string? error)
        {
            error = null;

            var list = new HashSet<Dw2DObject>();

            int controlY; ;
            foreach (var _object in _objects)
            {
                if (_object is Dw2DObject control)
                {
                    controlY = control.Y;
                    if (_bands is not null)
                    {
                        var band = _bands.FirstOrDefault((band) => band.Name == control.Band);
                        if (band is null)
                        {
                            error = $"Control {control.Name}' band {control.Band} not in band list";
                            return null;
                        }

                        controlY = control.Y + band.Position;

                    }
                    control.Y = controlY;
                    list.Add(control);
                }
            }

            return new DwControlMatrix(list, _bands);
        }
    }
}
