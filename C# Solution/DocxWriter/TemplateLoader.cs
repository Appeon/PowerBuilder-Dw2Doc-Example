using System.Text.Json;

namespace Appeon.DotnetDemo.DocumentWriter
{
    public class TemplateLoader
    {
        public static int LoadTemplates(string sourceFile, out TemplateListContainer? templates, out string? error)
        {
            templates = null;
            error = null;

            try
            {
                using var stream = new FileStream(sourceFile, FileMode.Open, FileAccess.Read);
                templates = new TemplateListContainer() { Templates = JsonSerializer.Deserialize<Template[]>(stream) };
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }

        public static int GetTemplateCount(in TemplateListContainer container)
        {
            return container.Templates?.Length ?? -1;
        }

        public static int GetTemplate(in TemplateListContainer container, int index, out Template? template, out string? error)
        {
            template = null;
            error = null;
            try
            {
                template = container.Templates?[index] ?? null;
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }
    }
}
