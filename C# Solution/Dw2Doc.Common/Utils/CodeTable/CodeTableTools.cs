using System.Text;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Utils.CodeTable;

public class CodeTableTools
{

    private const string RadioButtonSelected = "◉";
    private const string RadioButtonUnselected = "⊙";

    public static bool GetValueMap(string codetableString, out IDictionary<string, string>? codetable, out string? error)
    {
        error = null;
        var dict = new Dictionary<string, string>();

        try
        {
            string[] tokens = codetableString.Split("/", StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens)
            {
                var items = token.Split("\t");
                dict[items[1]] = items[0];
            }

        }
        catch (Exception e)
        {
            dict = null;
            error = e.Message;
        }
        codetable = dict;
        return error is null;
    }

    public static string[] GetDisplayValues(IDictionary<string, string> codemap) => codemap.Values.ToArray();

    public static string[] GetDataValues(IDictionary<string, string> codemap) => codemap.Keys.ToArray();

    public static string? BuildString(IDictionary<string, string> codetable, string? value, bool left, int columns)
    {
        var sb = new StringBuilder();

        int i = 0;
        foreach (var (data, display) in codetable)
        {
            if (!left)
            {
                sb.Append($"{GetButton(codetable, value, display)} ");
            }
            sb.Append($"{display}");

            if (left)
            {
                sb.Append($" {GetButton(codetable, value, display)}");
            }

            if ((i + 1) % columns == 0)
                sb.AppendLine();
            else
                sb.Append("    ");

            ++i;
        }

        sb.Length -= 2;

        return sb.ToString();
    }

    private static string? GetButton(IDictionary<string, string> codetable, string? value, string expected)
            => value is null
            ? RadioButtonUnselected
            : codetable.ContainsKey(value) && codetable[value] == expected
                ? RadioButtonSelected
                : RadioButtonUnselected;
}
