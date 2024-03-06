namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.Models
{
    public class AttributeTestResultCollection
    {
        public IList<AttributeTestResult> Results { get; set; }

        public AttributeTestResult Get(int index)
        {
            return Results[index];
        }

        public AttributeTestResultCollection(IList<AttributeTestResult> results)
        {
            Results = results;
        }
    }
}
