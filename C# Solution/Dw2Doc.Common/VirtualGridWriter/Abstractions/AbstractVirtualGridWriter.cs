using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects.DwObjectAttributes;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Models;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions
{
    public abstract class AbstractVirtualGridWriter
    {
        protected VirtualGrid.VirtualGrid VirtualGrid { get; }
        private IDictionary<string, DwObjectAttributesBase>? _previousDataSet0;
        protected bool IsClosed { get; set; }

        public IDictionary<string, DwObjectAttributesBase>? PreviousDataSet
        {
            get { return _previousDataSet0; }
            set { _previousDataSet0 = value; }
        }

        private readonly List<BandRows> _bandsWithoutRelatedHeaders;
        private readonly ISet<string> _unrepeatableBands;
        private int _currentBandIndex;
        private BandRows? _currentBandRow;
        private readonly Dictionary<string, bool> _bandsWithChanges = new();

        protected AbstractVirtualGridWriter(VirtualGrid.VirtualGrid virtualGrid)
        {
            _unrepeatableBands = new HashSet<string>(
                virtualGrid.BandRows
                    .Where(br => !br.IsRepeatable)
                    .Select(br => br.Name)
            );

            VirtualGrid = virtualGrid;
            _bandsWithoutRelatedHeaders = VirtualGrid.BandRows
                .Where(br => br.RelatedHeader is null)
                .ToList() ?? throw new InvalidOperationException("Virtual Grid has no Bands configured");
        }

        private int UpdateCurrentBand(int newBandIndex)
        {
            if (newBandIndex < 0)
                _currentBandRow = null;
            else
                _currentBandRow = _bandsWithoutRelatedHeaders[newBandIndex];
            _currentBandIndex = newBandIndex;
            return newBandIndex;
        }

        protected IList<ExportedCellBase>? ProcessNextLine(
            IDictionary<string, DwObjectAttributesBase>? dataSet,
            IList<string> bandsWithChanges)
        {
            // write headers

            // write groups (loop)

            var exportedCells = new List<ExportedCellBase>();


            _currentBandRow = _bandsWithoutRelatedHeaders[_currentBandIndex];
            if (dataSet is null)
            {
                while (_currentBandIndex >= 0)
                {
                    if (_bandsWithoutRelatedHeaders[_currentBandIndex].RelatedTrailers is not null)
                    {
                        foreach (var trailer in _bandsWithoutRelatedHeaders[_currentBandIndex].RelatedTrailers)
                        {
                            var exported = (WriteRows(trailer.Rows, PreviousDataSet));

                            if (exported is not null)
                            {
                                exportedCells.AddRange(exported);
                            }
                        }
                        UpdateCurrentBand(_currentBandIndex - 1);
                    }
                }
                return exportedCells;
            }

            if (bandsWithChanges.Count > 0 && _currentBandRow.Name != bandsWithChanges[0])
            {
                // We need this loop to repeat once more after the confition is met
                bool oneMore = false;
                while (_currentBandRow.Name != bandsWithChanges[0] || oneMore)
                {   // if dataset changed inlucing bands previous to the current one,
                    // write the trailers of the pending bands 
                    if (_currentBandRow.RelatedTrailers.Count > 0)
                    {
                        foreach (var trailer in _currentBandRow.RelatedTrailers)
                        {
                            var _exportedCells = WriteRows(trailer.Rows, PreviousDataSet);
                            if (_exportedCells is not null)
                                exportedCells.AddRange(_exportedCells);
                        }
                    }

                    // Guard against incorrectly looping past the condition
                    if (!oneMore)
                        UpdateCurrentBand(_currentBandIndex - 1);
                    if (_currentBandRow.Name == bandsWithChanges[0])
                        oneMore = !oneMore;
                }
            }

            for (; _currentBandIndex < _bandsWithoutRelatedHeaders.Count; ++_currentBandIndex)
            {
                var _exportedCells = WriteRows(_bandsWithoutRelatedHeaders[_currentBandIndex].Rows, dataSet);

                if (_exportedCells is not null) exportedCells.AddRange(_exportedCells);
            }
            --_currentBandIndex;

            return exportedCells;
        }

        public IList<ExportedCellBase>? EnterData(IDictionary<string, DwObjectAttributesBase>? dataSet, out string? error)
        {
            error = null;
            try
            {
                IList<ExportedCellBase>? exportedCells;
                if (IsClosed)
                {
                    throw new InvalidOperationException("Writer is already closed");
                }

                _bandsWithChanges.Clear();

                if (dataSet is not null)
                {
                    // Determine the bands that have changed
                    if (PreviousDataSet is null)
                    {
                        foreach (var band in VirtualGrid.BandRows)
                        {
                            _bandsWithChanges[band.Name] = band.Rows.Count > 0;
                        }

                    }
                    else
                    {

                        foreach (var (_, cellSet) in VirtualGrid.CellRepository.CellsByY)
                        {
                            foreach (var cell in cellSet)
                            {
                                if (!PreviousDataSet[cell.Object.Name]?.Equals(dataSet[cell.Object.Name]) ?? false)
                                {
                                    // Some columns pass different values per entry even though they're not supposed to change
                                    // for example, in the DW header there are CurrentTime()-like computed columns
                                    // these columns change each second, even though they display a static value
                                    // so we hack away this problem by excluding from the patch-generation controls
                                    // that belong to unupdatable bands (e.g. header)
                                    if (!_unrepeatableBands.Contains(cell.Object.Band))
                                        _bandsWithChanges[cell.Object.Band] = true;
                                }
                            }
                        }
                    }
                }

                exportedCells = ProcessNextLine(dataSet, _bandsWithChanges.Keys.ToList());
                PreviousDataSet = dataSet;

                return exportedCells;
            }
            catch (Exception e)
            {
                error = e.Message;
                return null;
            }
        }

        protected abstract IList<ExportedCellBase>? WriteRows(IList<RowDefinition> rows, IDictionary<string, DwObjectAttributesBase> data);

        public abstract bool Write(string? sheetname, out string? error);
    }
}
