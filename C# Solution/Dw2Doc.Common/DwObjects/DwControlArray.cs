namespace Appeon.DotnetDemo.Dw2Doc.Common.DwObjects
{
    public class DwControlArray
    {
        private readonly ICollection<DwObjectBase> _objects;
        private IList<DwBand>? _bands;

        public DwControlArray()
        {
            _objects = new HashSet<DwObjectBase>();
        }

        public DwControlArray Add2DControl(string name, string band, int x, int y, int width, int height)
        {
            var control = new Dw2DObject(name, band, x, y, width, height);

            _objects.Add(control);

            return this;
        }

        public DwControlArray Add1DControl(string name, string band, int x1, int x2, int y1, int y2)
        {
            var control = new Dw1DObject(name, band, x1, x2, y1, y2);

            _objects.Add(control);

            return this;
        }

        public DwControlArray SetBands(IList<DwBand> bands)
        {
            _bands = bands;
            return this;
        }


        public ICollection<Dw2DObject>? Build(out string? error)
        {
            error = null;

            var controls = new List<Dw2DObject>();

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

                    controls.Add(control);

                }
            }

            return controls;
        }
    }
}
