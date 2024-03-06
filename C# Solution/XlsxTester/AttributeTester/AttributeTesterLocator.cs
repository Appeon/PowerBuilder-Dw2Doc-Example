using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using System.Reflection;

namespace Appeon.DotnetDemo.Dw2Doc.XlsxTester.AttributeTester;

public class AttributeTesterLocator
{
    private static readonly IDictionary<Type, object> Cache;

    static AttributeTesterLocator()
    {
        Cache = new Dictionary<Type, object>();
    }

    public static object? Find(Type type)
    {
        if (Cache.ContainsKey(type))
        {
            return Cache[type];
        }

        var attributes = Assembly.GetExecutingAssembly()
            .GetTypes()
            .Where(t => t.GetCustomAttributes<TesterForAttribute>(true).Any())
            .SelectMany(t => t.GetCustomAttributes<TesterForAttribute>(true));

        foreach (var attrib in attributes)
        {
            if (Cache.ContainsKey(attrib.AttributeType))
            {
                continue;
            }
            var tester = Activator.CreateInstance(attrib.TesterType);
            if (tester is null)
            {
                return null;
            }
            Cache[attrib.AttributeType] = tester;
        }

        return (Cache.ContainsKey(type) ? Cache[type] : null);
    }

    public static object? Find(ExportedCellBase exportedCellBase)
    {
        return Find(exportedCellBase.AttributeType);
    }
}
