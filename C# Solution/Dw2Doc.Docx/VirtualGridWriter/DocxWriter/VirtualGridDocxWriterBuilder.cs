using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;

namespace Appeon.DotnetDemo.Dw2Doc.Docx.VirtualGridWriter.DocxWriter
{
    internal class VirtualGridDocxWriterBuilder : IVirtualGridWriterBuilder
    {
        public AbstractVirtualGridWriter? Build(VirtualGrid grid, out string? error)
        {
            error = null;
            return new VirtualGridDocxWriter(grid);
        }

        public AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid grid, string path, out string? error)
        {
            throw new NotImplementedException();
        }
    }
}
