using System.Collections;

namespace Appeon.PbExtensions.Collections;

public class Aggregators
{
    //public static IList<T> Merge<T>(ref IList<T> list1, IList<T> list2)
    //{
    //    foreach (var item in list2)
    //    {
    //        list1.Add(item);
    //    }

    //    return list1;
    //}

    public static IList Merge(ref IList list1, IList list2)
    {
        foreach (var item in list2)
        {
            list1.Add(item);
        }

        return list1;
    }
}
