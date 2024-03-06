using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System.Diagnostics.CodeAnalysis;

namespace Appeon.DotnetDemo.Dw2Doc.Docx.VirtualGridWriter.DocxWriter
{
    public class VirtualGridDocxWriter : AbstractVirtualGridWriter
    {
        private XWPFDocument _document;
        private bool _documentInitialized = false;
        private int _currentRow;

        private XWPFTable? _workingTable;

        public VirtualGridDocxWriter(VirtualGrid virtualGrid) : base(virtualGrid)
        {
            _document = new XWPFDocument();
        }

        [MemberNotNull(nameof(_workingTable))]
        private void InitDocument()
        {
#pragma warning disable CS8774 // Member must have a non-null value when exiting.
            if (_documentInitialized) return;
#pragma warning restore CS8774 // Member must have a non-null value when exiting.


            _documentInitialized = true;


            _workingTable = _document.CreateTable();

            _workingTable.GetCTTbl().tblPr.tblBorders.left.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.right.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.top.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.bottom.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.insideV.val = (ST_Border.none);
            _workingTable.GetCTTbl().tblPr.tblBorders.insideH.val = (ST_Border.none);


            for (int i = 0; i < VirtualGrid.Columns.Count; ++i)
            {
                _workingTable.AddNewCol();
            }


        }


        protected override void WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase> data)
        {
            InitDocument();

            if (data is null)
                return;

            foreach (var row in rows)
            {
                var newRow = _workingTable.CreateRow();
                newRow.Height = row.Size;

                foreach (var @object in row.Objects.Concat(row.FloatingObjects))
                {
                    XWPFTableCell cell;
                    if (@object.OwningColumn is not null)
                    {

                        cell = newRow.GetCell(@object.OwningColumn.IndexOffset);
                    }
                    else
                    { // floating cell
                        var floatingObject = @object as FloatingVirtualCell
                            ?? throw new InvalidOperationException("Cell has no owner column and is not floating");
                    }



                }
            }

        }


        public override bool Write(string path, out string? error)
        {
            error = null;

            if (path is null)
            {
                error = "No file specified";
                return false;
            }

            try
            {
                using var stream = File.Create(path);

                _document.Write(stream);
            }
            catch (IOException e)
            {
                error = e.ToString();
                return false;
            }

            return true;
        }
    }
}
