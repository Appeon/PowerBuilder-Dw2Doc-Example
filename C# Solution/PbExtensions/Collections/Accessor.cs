using System.Collections;

namespace Appeon.PbExtensions.Collections;

public class Accessor
{
    public static object Get(IList list, int index) => list[index];
}
