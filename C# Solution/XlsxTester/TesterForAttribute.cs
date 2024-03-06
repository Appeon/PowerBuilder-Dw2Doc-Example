namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class TesterForAttribute : Attribute
    {
        public Type TesterType { get; set; }
        public Type AttributeType { get; set; }

        public TesterForAttribute(Type attribute, Type tester)
        {
            TesterType = tester;
            AttributeType = attribute;
        }
    }
}
