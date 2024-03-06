using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;

namespace Appeon.DotnetDemo.Dw2Doc.Common.Exceptions
{
    /// <summary>
    /// Exception that occurs when the objects in a <see cref="VirtualGrid.VirtualGrid"/>'s rows and
    /// columns aren't consistent.
    /// </summary>
    internal class GridObjectCountInconsistencyException : Exception
    {
        public string? ObjectName { get; set; }

        public override string Message => ObjectName is null ? base.Message : $"Inconsistent control: {ObjectName}";

        public GridObjectCountInconsistencyException()
        {
        }

        public GridObjectCountInconsistencyException(DwObjectBase @object)
        {
            ObjectName = @object.Name;
        }
    }
}
