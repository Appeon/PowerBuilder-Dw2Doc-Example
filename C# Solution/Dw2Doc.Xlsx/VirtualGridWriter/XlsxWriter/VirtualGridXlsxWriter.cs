using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Extensions;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.Helpers;
using Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.Renderers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.XlsxWriter
{
    public class VirtualGridXlsxWriter : AbstractVirtualGridWriter
    {
        private int _startRowOffset;
        private int _startColumnOffset;
        private int _currentRowOffset;
        private readonly XSSFWorkbook _workbook;
        private readonly XSSFSheet _sheet;
        private readonly XSSFDrawing _drawingPatriarch;
        private readonly ISet<ColumnDefinition> _resizedColumns;
        private readonly IDictionary<string, int> _pictureCache;
        private bool _writerInitialized = false;
        private bool _appendingToSheet = false;
        private string _path;
        private bool _closed = false;

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path)
            : base(grid)
        {
            _workbook = new XSSFWorkbook();
            _sheet = (XSSFSheet)_workbook.CreateSheet();
            _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
            _resizedColumns = new HashSet<ColumnDefinition>();

            _resizedColumns.AddRange(VirtualGrid.Columns);
            _pictureCache = new Dictionary<string, int>();
            _path = path;
        }

        internal VirtualGridXlsxWriter(VirtualGrid grid, string path, string sheetname) : base(grid)
        {
            _path = path;

            _workbook = new XSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read));

            if (_workbook.GetSheet(sheetname) is not null)
            {
                throw new ArgumentException("Specified sheet name already exists", nameof(sheetname));
            }

            _sheet = (XSSFSheet)_workbook.CreateSheet(sheetname);
            _workbook.SetActiveSheet(_workbook.GetSheetIndex(_sheet));

            _drawingPatriarch = (XSSFDrawing)_sheet.CreateDrawingPatriarch();
            _resizedColumns = new HashSet<ColumnDefinition>();

            _resizedColumns.AddRange(VirtualGrid.Columns);
            _pictureCache = new Dictionary<string, int>();

            _appendingToSheet = true;
        }

        private void InitWriter()
        {
            if (_writerInitialized) return;

            var row = _sheet.CreateRow(_startRowOffset);
            foreach (var column in _resizedColumns)
            {
                row.CreateCell(column.IndexOffset + _startColumnOffset);
                _sheet.SetColumnWidth(column.IndexOffset + _startColumnOffset, (int)UnitConversion.PixelsToColumnWidth(column.Size));
            }

            _sheet.RemoveRow(row);

            _writerInitialized = true;
        }

        protected override IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase>? data)
        {
            InitWriter();
            if (data is null)
                return null;

            if (_closed)
            {
                throw new InvalidOperationException("This writer has already been closed");
            }

            var exportedCells = new List<ExportedCellBase>();
            ExportedCellBase? exportedCell = null;

            foreach (var row in rows)
            {
                var xRow = _sheet.CreateRow(_startRowOffset + _currentRowOffset);

                checked
                {
                    xRow.Height = (short)(row.Size).PixelsToTwips(Windows.ScreenTools.MeasureDirection.Vertical);
                    //xRow.Height = (short)row.Size;
                }

                ICell? previousCell = null;
                int lastOccupiedColumn = 0;
                foreach (var cell in row.Objects.Concat(row.FloatingObjects))
                {
                    var attribute = data[cell.Object.Name];

                    if (!attribute.IsVisible)
                    {
                        continue;
                    }
                    var renderer = RendererLocator.Find(attribute.GetType())
                        ?? throw new InvalidOperationException($"Could not find renderer for attribute [{attribute.GetType().FullName}]");

                    if (renderer is XlsxPictureRenderer pictureRenderer)
                    {
                        pictureRenderer.SetPictureCache(_pictureCache);
                    }

                    if (cell.OwningColumn is not null && !cell.Object.Floating)
                    { // cell is not floating
                        var xCell = xRow.CreateCell(cell.OwningColumn.IndexOffset + _startColumnOffset);

                        var style = _workbook.CreateCellStyle();
                        switch (attribute)
                        {
                            case DwTextAttributes txt:
                                //style.Alignment = txt.Alignment.ToNpoiHorizontalAlignment();
                                //xCell.CellStyle = style;
                                break;
                        }

                        if (xCell.ColumnIndex - lastOccupiedColumn > 2)
                        {
                            var cellRange = new CellRangeAddress(
                                xCell.RowIndex,
                                xCell.RowIndex,
                                lastOccupiedColumn + 1,
                                xCell.ColumnIndex - 1);

                            _sheet.AddMergedRegion(cellRange);
                        }

                        lastOccupiedColumn = cell.OwningColumn.IndexOffset + _startColumnOffset;

                        previousCell = xCell;
                        exportedCell = (renderer.Render(_sheet, cell, attribute, xCell));

                        if (exportedCell is not null)
                        {
                            exportedCells.Add(exportedCell);
                        }

                        lastOccupiedColumn = cell.OwningColumn.IndexOffset + _startColumnOffset;
                        if (cell.ColumnSpan > 1)
                        {
                            var cellRange = new CellRangeAddress(
                                xCell.RowIndex,
                                xCell.RowIndex,
                                xCell.ColumnIndex,
                                xCell.ColumnIndex + cell.ColumnSpan - 1);

                            lastOccupiedColumn += (cell.ColumnSpan - 1);

                            _sheet.AddMergedRegion(cellRange);
                        }

                    }
                    else
                    { // cell is floating
                        if (cell is not FloatingVirtualCell floatingCell)
                            throw new InvalidOperationException("Non-floating cell in FloatingObjects list");
                        previousCell = null;
                        exportedCell = renderer.Render(_sheet, floatingCell,
                            data[cell.Object.Name],
                            (_startColumnOffset + floatingCell.Offset.StartColumn.IndexOffset,
                                _currentRowOffset,
                                _drawingPatriarch));

                        if (exportedCell is not null)
                        {
                            exportedCells.Add(exportedCell);
                        }
                    }
                }


                int unoccupiedTrailingColumns = 0;
                if (row.IsFiller
                    && VirtualGrid.Columns.Count > 1
                    && row.Objects.Count == 0
                    && row.FloatingObjects.Count == 0
                    || (unoccupiedTrailingColumns = VirtualGrid.Columns.Count - lastOccupiedColumn) > 2)
                {
                    var cellRange = new CellRangeAddress(
                        _currentRowOffset + _startRowOffset,
                        _currentRowOffset + _startRowOffset,
                        _startColumnOffset + unoccupiedTrailingColumns > 2 ? (lastOccupiedColumn + 1) : 0,
                        _startColumnOffset + VirtualGrid.Columns.Count - 1);

                    _sheet.AddMergedRegion(cellRange);
                }

                ++_currentRowOffset;
            }

            return exportedCells;
        }

        public override bool Write(string? sheetname, out string? error)
        {
            error = null;

            if (_path is null)
            {
                error = "No file specified";
                return false;
            }
            try
            {
                using var stream = File.Exists(_path)
                    ? new FileStream(_path, FileMode.Truncate, FileAccess.Write)
                    : new FileStream(_path, FileMode.Create, FileAccess.Write);
                _workbook.Write(stream);

                _workbook.Close();
                _closed = true;

            }
            catch (IOException e)
            {
                error = e.Message;
                return false;
            }

            return true;
        }
    }
}
