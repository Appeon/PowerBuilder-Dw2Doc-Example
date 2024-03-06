using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;

namespace Appeon.DotnetDemo.Dw2Doc.DwData.DataSet
{
    public class DataSetBuilder
    {
        private IDictionary<string, DwObjectAttributesBase> _attributes;

        public DataSetBuilder()
        {
            _attributes = new Dictionary<string, DwObjectAttributesBase>();
        }

        public void AddControlAttribute(string controlName, DwObjectAttributesBase attribute)
        {
            _attributes[controlName] = attribute;
        }

        public void Clear() => _attributes = new Dictionary<string, DwObjectAttributesBase>();

        public IDictionary<string, DwObjectAttributesBase> Build() => _attributes;
    }
}
