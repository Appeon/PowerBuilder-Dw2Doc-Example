namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions
{
    public interface IVirtualGridWriterBuilder
    {
        AbstractVirtualGridWriter? Build(VirtualGrid.VirtualGrid grid, out string? error);

        AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid.VirtualGrid grid, string path, out string? error);
    }
}
