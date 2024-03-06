using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Extensions
{
    public static class EntityExtensions
    {
        internal static IList<T> ChainToList<T>(this T source)
            where T : EntityDefinition<T>
        {
            return ChainIntoList(source, new List<T>());
        }

        internal static IList<T> ChainIntoList<T>(this T source, IList<T> list)
            where T : EntityDefinition<T>
        {
            T? start = source;
            while (start.PreviousEntity is not null)
            {
                start = start.PreviousEntity;
            }

            while (start is not null)
            {
                list.Add(start);
                start = start.NextEntity;
            }

            return list;
        }
    }
}
