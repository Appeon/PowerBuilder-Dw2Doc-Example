using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Attributes;
using System.Reflection;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers
{
    internal static class RendererLocator
    {
        private static readonly IDictionary<Type, object> Cache;

        static RendererLocator()
        {
            Cache = new Dictionary<Type, object>();
        }

        public static AbstractXlsxRenderer? Find(Type type)
        {
            if (Cache.ContainsKey(type))
            {
                return (AbstractXlsxRenderer)Cache[type];
            }

            var attributes = Assembly.GetExecutingAssembly()
                .GetTypes()
                .Where(t => t.GetCustomAttributes<RendererForAttribute>(true).Any())
                .SelectMany(t => t.GetCustomAttributes<RendererForAttribute>(true));

            foreach (var attrib in attributes)
            {
                if (Cache.ContainsKey(attrib.TargetType))
                {
                    continue;
                }
                var @obj = Activator.CreateInstance(attrib.OwningType);
                if (@obj is null)
                {
                    return null;
                }
                Cache[attrib.TargetType] = @obj;
            }

            return (Cache.ContainsKey(type) ? (AbstractXlsxRenderer)Cache[type] : null);
        }
    }
}
