namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes
{
    [AttributeUsage(AttributeTargets.Class, Inherited = true, AllowMultiple = false)]
    public sealed class RendererForAttribute : Attribute
    {
        public Type TargetType { get; }
        public Type OwningType { get; }

        public RendererForAttribute(Type targetType, Type owningType)
        {
            TargetType = targetType;
            OwningType = owningType;
        }
    }
}
