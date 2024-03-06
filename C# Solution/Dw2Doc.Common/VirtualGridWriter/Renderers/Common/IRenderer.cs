using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Renderers.Common
{
    public interface IRenderer<TContext, TCell, TAttrib, TTarget>
        where TCell : VirtualCell
    {
        /// <summary>
        /// Render the object
        /// </summary>
        ExportedCellBase? Render(TContext context, TCell cell, TAttrib attribute, TTarget renderTarget);
    }
}
