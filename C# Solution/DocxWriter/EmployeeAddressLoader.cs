using System.Text.Json;

namespace Appeon.DotnetDemo.DocumentWriter
{
    public class EmployeeAddressLoader
    {
        public static int Load(string path, out object? repo, out string? error)
        {
            repo = null;
            error = null;

            try
            {
                using var stream = File.OpenRead(path);
                repo = JsonSerializer.Deserialize<IList<EmployeeAddress>>(stream);
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
            return 1;
        }

        public static int SearchById(int id, in object repo, out object? entry, out string? error)
        {
            entry = null;
            error = null;

            try
            {
                var _repo = (IList<EmployeeAddress>)repo;
                if (_repo is null)
                {
                    error = "Invalid repo";
                    return -1;
                }

                entry = _repo.Where(a => a.BusinessEntityId == id)
                    .FirstOrDefault();
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }

            return 1;
        }
    }
}
