using Appeon.DotnetDemo.Dw2Doc.Common.DwObjects;
using Appeon.DotnetDemo.Dw2Doc.Common.Enums;
using Appeon.DotnetDemo.Dw2Doc.Common.Exceptions;
using Appeon.DotnetDemo.Dw2Doc.Common.Extensions;
using System.Drawing;

namespace Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid
{
    public class VirtualGridBuilder
    {
        public int XThreshold { get; set; } = 3;
        public int YThreshold { get; set; } = 3;

        public VirtualGrid? Build(
            DwControlMatrix controlMatrix,
            byte dwProcessing,
            bool simplifyBands,
            bool discardFloating,
            out string? error)
        {
            /// TODO: take into account the band's height when adding padding to the rows
            error = null;
            try
            {
                var dwType = (DwType)dwProcessing;
                //VirtualGrid grid = new(DefineRows(controlMatrix), DefineColumns(controlMatrix));

                var cellRepo = controlMatrix.Controls.ToVirtualCellRepository();

                var (virtualRows,
                    rowsPerBand,
                    floatingControls) = DefineRows(cellRepo.CellsByY,
                    controlMatrix.Bands,
                    simplifyBands);



                foreach (var row in virtualRows)
                {
                    int i = 0;
                    var controlsToRemoveFromRow = new HashSet<VirtualCell>();

                    /// look for objects in the same row with overlapping X, and convert one of them to 
                    /// a floating control
                    foreach (var control in row.Objects)
                    {
                        for (int j = i; j < row.Objects.Count; ++j)
                        {
                            if (row.Objects[j] != control
                                && (
                                    (row.Objects[j].X >= control.X && row.Objects[j].X < control.RightBound)
                                    || (row.Objects[j].RightBound > control.X && row.Objects[j].RightBound <= control.RightBound)
                                ))
                            {
                                floatingControls.Add(row.Objects[j]);
                                controlsToRemoveFromRow.Add(row.Objects[j]);
                            }
                        }
                        i++;
                    }

                    foreach (var control in controlsToRemoveFromRow)
                    {
                        control.OwningRow = null;
                        control.OwningColumn = null;
                        row.Objects.Remove(control);
                    }
                }

                var virtualColumns = DefineColumns(cellRepo.CellsByX, floatingControls);

                switch (dwType)
                {
                    case DwType.Grid:
                        {
                            virtualColumns = RemoveFillerColumns(virtualColumns);
                            break;
                        }
                }

                if (discardFloating)
                {
                    floatingControls = floatingControls
                        .Where(x => x.Object.Floating)
                        .ToHashSet();
                }

                // validate integrity (rows and columns must contain the exact same controls)
                CheckIntegrity(virtualRows, virtualColumns);

                /// Add padding entities to allow for placing floating controls that overlap existing entities
                AddNecessaryPaddingColumns(ref virtualColumns, floatingControls);

                /// Make sure the band has the same height as it defined in the DW
                /// if not, insert rows to match
                int bottomMostBound;
                int missingPadding;
                RowDefinition? bottomMostRow = null;
                foreach (var band in controlMatrix.Bands)
                {
                    bottomMostBound = int.MinValue;
                    if (rowsPerBand.ContainsKey(band.Name))
                    {
                        foreach (var row in rowsPerBand[band.Name])
                        {
                            if (bottomMostBound < row.Bound)
                            {
                                bottomMostBound = row.Bound;
                                bottomMostRow = row;
                            }
                        }

                        if (bottomMostRow is not null)
                        {

                            missingPadding = band.Bound - bottomMostBound;
                            if (missingPadding > 0)
                            {
                                var newRow = new RowDefinition
                                {
                                    BandName = band.Name,
                                    IsFiller = true,
                                    IsPadding = true,
                                    Size = missingPadding,
                                    PreviousEntity = bottomMostRow,
                                    NextEntity = bottomMostRow.NextEntity
                                };
                                if (bottomMostRow.NextEntity is not null)
                                {
                                    bottomMostRow.NextEntity.PreviousEntity = newRow;
                                }

                                bottomMostRow.NextEntity = newRow;

                                virtualRows.Insert(virtualRows.IndexOf(bottomMostRow) + 1, newRow);
                                rowsPerBand[band.Name].Add(newRow);
                            }

                        }
                    }
                }



                var paddingRow = AddNecessaryPaddingRows(ref virtualRows, floatingControls);
                if (paddingRow is not null)
                {
                    var lastRow = paddingRow.PreviousEntity;
                    foreach (var (band, rows) in rowsPerBand)
                    {
                        if (rows.Contains(lastRow))
                        {
                            rowsPerBand[band].Add(paddingRow);
                            break;
                        }
                    }
                }

                /// dissolve filler rows
                /// 

                var fillerRows = virtualRows.Where(row => row.IsFiller)
                    .ToList();

                foreach (var row in fillerRows)
                {

                    bool dissolved = false;
                    if (row.Objects.Count > 0 || row.FloatingObjects.Count > 0)
                    {
                        throw new Exception("Filler row has content");
                    }

                    if (row.PreviousEntity is not null)
                    {

                        // Compensate floating controls' height with the height of the row we're about to remove
                        foreach (var floatingControl in floatingControls)
                        {
                            if (floatingControl.LowerBound >= row.PreviousEntity.Offset && floatingControl.LowerBound <= row.PreviousEntity.Bound)
                            {
                                floatingControl.Height += row.Size;
                            }
                        }

                        row.PreviousEntity.Size += row.Size;
                        row.PreviousEntity.NextEntity = row.NextEntity;



                        dissolved = true;
                    }

                    if (row.NextEntity is not null)
                    {
                        row.NextEntity.PreviousEntity = row.PreviousEntity;
                        if (!dissolved)
                        {
                            // Compensate floating controls' height and Y with the height of the row we're about to remove
                            foreach (var floatingControl in floatingControls)
                            {
                                if (floatingControl.Y >= row.NextEntity.Offset && floatingControl.Y <= row.NextEntity.Bound)
                                {
                                    floatingControl.Height += row.Size;
                                    floatingControl.Y -= row.Size;
                                }
                            }

                            row.NextEntity.Size += row.Size;
                        }
                    }

                    rowsPerBand[row.BandName].Remove(row);
                    virtualRows.Remove(row);
                }

                // assign floating objects to their rows

                FloatingCellOffset? cellOffset;
                foreach (var floatingControl in floatingControls)
                {
                    cellOffset = MapFloatingControl(virtualColumns,
                        virtualRows,
                        floatingControl);
                    if (cellOffset is null)
                    {
                        throw new InvalidOperationException("Could not find an appropriate reference column for cell");
                    }


                    //floatingControl
                    //    .OwningRow!
                    cellOffset
                        .StartRow
                        .FloatingObjects
                        .Add(FloatingVirtualCell.FromVirtualCell(
                                floatingControl,
                                cellOffset));
                }


                VirtualGrid grid = new(virtualRows,
                    virtualColumns,
                    controlMatrix.Bands.DwBandsToBandRows(rowsPerBand),
                    cellRepo,
                    dwType);


                return grid;
            }
            catch (Exception e)
            {
                error = e.Message;
                return null;
                throw;
            }
        }



        private static void CheckIntegrity(
            ICollection<RowDefinition> rows,
            ICollection<ColumnDefinition> columns)
        {
            var map = new Dictionary<VirtualCell, byte>();

            foreach (var row in rows)
            {
                // Not taking into account floating objects
                foreach (var control in row.Objects)
                {
                    if (!map.ContainsKey(control))
                        map[control] = 0;
                    map[control]++;
                }
            }

            // All controls in the columns must exist as solid cells in the rows, 
            foreach (var col in columns)
            {
                foreach (var control in col.Objects)
                {
                    if (!map.ContainsKey(control))
                        map[control] = 0;
                    map[control]++;
                }
            }

            foreach (var (control, occurrences) in map)
            {
                if (occurrences != 2)
                {
                    throw new GridObjectCountInconsistencyException(control.Object);
                }
            }
        }

        private static IList<ColumnDefinition> RemoveFillerColumns(IList<ColumnDefinition> columns)
        {
            var newColumnList = new List<ColumnDefinition>();

            foreach (var column in columns)
            {
                if (column.IsFiller && column.Objects.Count == 0)
                {
                    column.RemoveFromChain();
                }
                else
                {
                    column.CalculateOffset();
                    newColumnList.Add(column);
                }
            }

            return newColumnList;

        }

        private static ColumnDefinition? AddNecessaryPaddingColumns(ref IList<ColumnDefinition> columns, ISet<VirtualCell> cells)
        {
            return AddNecessaryPaddingEntities(ref columns, cells, (cell) => cell.RightBound);
        }

        private static RowDefinition? AddNecessaryPaddingRows(ref IList<RowDefinition> rows, ISet<VirtualCell> cells)
        {
            var row = AddNecessaryPaddingEntities(ref rows, cells, (cell) => cell.LowerBound);
            if (row is not null)
                row.BandName = rows.Last().BandName;
            return row;
        }

        private static T? AddNecessaryPaddingEntities<T>(ref IList<T> entities,
                ISet<VirtualCell> cells,
                Func<VirtualCell, int> boundAccessor)
            where T : EntityDefinition<T>, new()
        {
            int furthestReachingControlBound = int.MinValue;
            int cellBound;
            foreach (var cell in cells)
            {
                cellBound = boundAccessor(cell);
                furthestReachingControlBound = cellBound > furthestReachingControlBound ? cellBound : furthestReachingControlBound;
            }

            var lastColumn = entities.Last();

            if (furthestReachingControlBound > lastColumn.Bound)
            {
                var newEntity = new T
                {
                    PreviousEntity = lastColumn,
                    Size = furthestReachingControlBound - lastColumn.Bound,
                };

                lastColumn.NextEntity = newEntity;
                entities.Add(newEntity);
                return newEntity;
            }

            return null;
        }

        private (
            IList<RowDefinition>, // List of Rows
            IDictionary<string, IList<RowDefinition>>, // band/row map
            ISet<VirtualCell> // floating control set,
            )
            DefineRows(SortedList<int, IList<VirtualCell>> controls,
                        IList<DwBand> bands,
                        bool simplifyBand
            )
        {
            // get overlapping chain
            // find first control that doesn't overlap the first one of the chain
            //  this will create 2 rows: one with the first component and the in-between (floating), and
            //      the other one. This will become a new chain including the remaining controls in the chain


            RowDefinition? currentRow = null;
            RowDefinition? previousRow = null;

            var rowsPerBand = new Dictionary<string, IList<RowDefinition>>();
            var rowsInCurrentBand = new List<RowDefinition>();
            var floatingControlSet = new HashSet<VirtualCell>();

            string? previousBand = null;
            string? currentBand = null;

            foreach (var (y, controlSet) in controls)
            {// create a row that pads the initial control offset, exluding floating controls
                var firstNonFloatingCell = controlSet.Where(c => !c.Object.Floating)
                    .FirstOrDefault();


                if (firstNonFloatingCell is not null)
                {
                    if (y > YThreshold)
                    {
                        previousRow = new RowDefinition()
                        {
                            Size = y,
                            BandName = firstNonFloatingCell.Object.Band,
                            IsPadding = true,
                        };

                        rowsInCurrentBand.Add(previousRow);
                    }

                    break;
                }
            }

            foreach (var (y, controlSet) in controls)
            // each controlSet contains controls that have same Y
            {
                // Get smallest control and define row around it, rest of controls will be floating 
                // TODO: design better algorithm to make the most controls be not floating
                int smallestSize = int.MaxValue;

                if (controlSet.Count == 0)
                    throw new InvalidOperationException("Cannot calculate rows from no controls");

                VirtualCell? smallestControl = null;

                var controlsNotConfiguredAsFloating = new HashSet<VirtualCell>();
                var controlsConfiguredAsFloating = new HashSet<VirtualCell>();


                foreach (var control in controlSet)
                {
                    if (control.Object.Floating)
                    {
                        controlsConfiguredAsFloating.Add(control);

                        continue;
                    }
                    controlsNotConfiguredAsFloating.Add(control);

                    if (control.Height < smallestSize)
                    {
                        smallestSize = control.Height;
                        smallestControl = control;
                    }
                }
                floatingControlSet.AddRange(controlsConfiguredAsFloating);

                if (smallestControl is null && controlsNotConfiguredAsFloating.Count > 0)
                {
                    throw new NullReferenceException("Could not determine the smallest control in the set");
                }


                var equal = new List<VirtualCell>();
                var larger = new List<VirtualCell>();

                string? newBand = null;

                // determine which controls are same height and which are larger
                foreach (var control in controlsNotConfiguredAsFloating)
                {
                    if (control.Height == smallestControl.Height)
                        equal.Add(control);
                    else
                    {
                        //control.Y -= ySkew;
                        larger.Add(control);
                    }
                    if (newBand is not null && newBand != control.Object.Band)
                    {
                        throw new DwBandException(control.Object.Name, newBand, control.Object.Band);
                    }
                    newBand = control.Object.Band;
                }

                if (newBand is null)
                {
                    if (controlsConfiguredAsFloating.Count == 0)
                    {
                        throw new InvalidOperationException("No controls to process. This should not happen.");
                    }

                    newBand = controlsConfiguredAsFloating.First().Object.Band;
                }

                currentBand = newBand;

                if (previousBand is not null && currentBand != previousBand)
                {
                    rowsPerBand[previousBand] = rowsInCurrentBand;


                    /// Create empty row to fill the rest of the band's space
                    int bandHeight = bands.Where(b => b.Name == previousBand).Single().Bound;
                    int rowBandDiff = bandHeight - rowsPerBand[previousBand].LastOrDefault()?.Bound ?? bandHeight;
                    if (rowBandDiff > YThreshold)
                    {
                        currentRow = new RowDefinition()
                        {
                            Size = rowBandDiff,
                            PreviousEntity = previousRow,
                            IsFiller = true,
                            BandName = previousBand,
                        };

                        if (previousRow is not null)
                        {
                            previousRow.NextEntity = currentRow;
                        }

                        rowsInCurrentBand.Add(currentRow);
                    }
                    if (simplifyBand && SimplifyBand(previousBand, rowsInCurrentBand))
                    {
                        previousRow = rowsInCurrentBand[rowsInCurrentBand.Count - 1];
                    }
                    previousBand = currentBand;
                    rowsInCurrentBand = new List<RowDefinition>();
                }

                bool repeat;
                if (controlsNotConfiguredAsFloating.Any()) do
                    { // this loop saves code when inserting a blank row
                        repeat = false;

                        // difference between previous row and current control set
                        int distanceToPreviousRowEnd = (y - previousRow?.Bound ?? 0);
                        int distanceToPreviousRowStart = (y - previousRow?.Offset ?? 0);
                        int rowHeightModifier = 0;
                        if (previousRow is null || (distanceToPreviousRowEnd >= -YThreshold && distanceToPreviousRowEnd <= YThreshold))
                        {
                            rowHeightModifier = (int)Math.Ceiling(distanceToPreviousRowEnd * 0.5);

                            if (previousRow is not null)
                            {
                                previousRow.Size += (int)Math.Floor(distanceToPreviousRowEnd * 0.5);
                                foreach (var obj in previousRow.Objects)
                                {
                                    obj.Height += rowHeightModifier;
                                }
                                previousRow.CompensatingSkew = true;
                            }


                            currentRow = new RowDefinition()
                            {
                                PreviousEntity = previousRow,
                                Size = smallestControl.Height + rowHeightModifier,
                                CompensatingSkew = rowHeightModifier != 0,
                                Objects = equal,
                                //FloatingObjects = larger,
                                BandName = currentBand,
                            };

                            equal.ForEach(obj =>
                            {
                                obj.Y -= rowHeightModifier;
                                obj.Height += rowHeightModifier;
                            });

                            equal.AssignToEntity(currentRow);
                            larger.AssignToEntity(currentRow);
                            floatingControlSet.AddRange(larger);
                            rowsInCurrentBand.Add(currentRow);
                        }
                        else if (distanceToPreviousRowEnd < 0)
                        { // if Y is inside the bounds of previous control, make the set float in the previous row
                            currentRow = previousRow;
                            //currentRow.FloatingObjects.AddRange(controlSet);
                            var controlsToAddToPreviousRow = controlsNotConfiguredAsFloating.Where(
                                c => c.Height == previousRow.Size
                                                && distanceToPreviousRowStart < YThreshold)
                                .ToList();
                            controlsToAddToPreviousRow.AssignToEntity(previousRow);
                            previousRow.Objects.AddRange(controlsToAddToPreviousRow);
                            floatingControlSet.AddRange(controlsNotConfiguredAsFloating.Except(controlsToAddToPreviousRow));
                            equal.AssignToEntity(currentRow);
                            larger.AssignToEntity(currentRow);

                            previousRow = currentRow.PreviousEntity;
                        }
                        else
                        { // controlset is too far from the previous row, create intermediate row
                            currentRow = new RowDefinition
                            {
                                PreviousEntity = previousRow,
                                Size = distanceToPreviousRowEnd,
                                BandName = currentBand,
                                IsFiller = true,
                            };
                            rowsInCurrentBand.Add(currentRow);
                            repeat = true;
                        }
                        if (previousRow is not null)
                            previousRow.NextEntity = currentRow;


                        currentRow.SortControlsByX();
                        previousRow = currentRow;

                    } while (repeat);

                previousBand = currentBand;
            }



            if (previousBand is not null)
            {
                rowsPerBand[previousBand] = rowsInCurrentBand;

                int bandHeight = bands.Where(b => b.Name == previousBand).Single().Bound;
                int rowBandDiff = bandHeight - rowsPerBand[previousBand].LastOrDefault()?.Bound ?? bandHeight;
                if (rowBandDiff > YThreshold)
                {
                    currentRow = new RowDefinition()
                    {
                        Size = rowBandDiff,
                        PreviousEntity = previousRow,
                        IsFiller = true,
                        BandName = previousBand,
                    };

                    if (previousRow is not null)
                    {
                        previousRow.NextEntity = currentRow;
                    }

                    rowsInCurrentBand.Add(currentRow);
                }
                if (simplifyBand && SimplifyBand(previousBand, rowsInCurrentBand))
                {
                    previousRow = rowsInCurrentBand[rowsInCurrentBand.Count - 1];
                }
            }



            return (currentRow?.ChainToList() ?? new List<RowDefinition>(), rowsPerBand, floatingControlSet);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="band"></param>
        /// <param name="rowsInBand"></param>
        /// <returns>true if user needs to update the last row reference</returns>
        private static bool SimplifyBand(string band, IList<RowDefinition> rowsInBand)
        {

            if ((band == "detail" || band == "header")
                && rowsInBand.Count == 3
                && rowsInBand[0].IsFiller
                && rowsInBand[2].IsFiller
                && !rowsInBand[0].IsPadding
                && !rowsInBand[2].IsPadding)
            {
                rowsInBand[1].Size += (rowsInBand[0].Size + rowsInBand[2].Size);
                rowsInBand[1].PreviousEntity = rowsInBand[0].PreviousEntity;
                if (rowsInBand[0].PreviousEntity is not null)
                {
                    rowsInBand[0].PreviousEntity!.NextEntity = rowsInBand[1];
                }

                rowsInBand[1].NextEntity = rowsInBand[2].NextEntity;
                if (rowsInBand[2].NextEntity is not null)
                {
                    rowsInBand[2].NextEntity!.PreviousEntity = rowsInBand[1];
                }

                rowsInBand.Remove(rowsInBand[2]);
                rowsInBand.Remove(rowsInBand[0]);
                return true;
            }
            if ((band == "detail" || band == "header")
                && rowsInBand.Count == 2
                && !rowsInBand[0].IsPadding
                && !rowsInBand[1].IsPadding)
            {

                if (rowsInBand[0].IsFiller)
                {
                    rowsInBand[1].Size += (rowsInBand[0].Size);
                    if (rowsInBand[0].PreviousEntity is not null)
                    {
                        rowsInBand[0].PreviousEntity!.NextEntity = rowsInBand[1];
                    }
                    rowsInBand[1].PreviousEntity = rowsInBand[0].PreviousEntity;
                    rowsInBand.Remove(rowsInBand[0]);
                    return false;
                }
                else
                {
                    rowsInBand[0].Size += (rowsInBand[1].Size);
                    if (rowsInBand[1].NextEntity is not null)
                    {
                        rowsInBand[1].NextEntity!.PreviousEntity = rowsInBand[0];
                    }
                    rowsInBand[0].NextEntity = rowsInBand[1].NextEntity;
                    rowsInBand.Remove(rowsInBand[1]);
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Define the rows 
        /// </summary>
        /// <param name="controls">List of controls ordered by X. Items might be removed from this list if they overlap with other
        /// controls on Y</param>
        /// <param name="floatingControls">List of floating controls that will be excluded from calculations. Items might
        /// be added to this list if they overlap with other controls on Y</param>
        /// <returns></returns>
        private IList<ColumnDefinition> DefineColumns(
                SortedList<int, ISet<VirtualCell>> controls,
                ISet<VirtualCell> floatingControls
            )
        {

            var columns = new List<ColumnDefinition>();
            ColumnDefinition? workingColumn;

            foreach (var (x, controlSet) in controls)
            {
                foreach (var control in controlSet)
                {
                    // skip floating controls
                    if (floatingControls.Contains(control))
                        continue;


                    var overlappingColumns = columns.Count > 0 ? GetOverlappingColumns(columns.First(), control) : null;
                    if (overlappingColumns is null || overlappingColumns.Columns.Count == 0)
                    { // No overlapping columns, object must be separated 
                        bool repeat;
                        do
                        {
                            repeat = false;

                            var previousColumn = columns.LastOrDefault();
                            int dx = control.X - (previousColumn?.Bound ?? 0);
                            workingColumn = new ColumnDefinition()
                            {
                                Size = dx
                            };

                            if (dx <= XThreshold)
                            {
                                workingColumn.Objects = new List<VirtualCell>() { control };
                                workingColumn.Size = control.Width + dx;
                                control.OwningColumn = workingColumn;
                            }
                            else
                            {
                                repeat = true;
                                workingColumn.IsFiller = true;
                            }
                            workingColumn.PreviousEntity = previousColumn;
                            if (previousColumn is not null)
                            {
                                previousColumn.NextEntity = workingColumn;
                            }

                            columns.Add(workingColumn);

                        } while (repeat);
                    }
                    else
                    { // overlapping columns

                        var firstOverlappingColumn = overlappingColumns
                                .Columns
                                .First();

                        var lastOverlappingColumn = overlappingColumns
                                .Columns
                                .Last();

                        // we will need to correct the column span to accomodate for new columns
                        int additionalSpan = 0;

                        if (overlappingColumns.OffsetRight > XThreshold)
                        { // control terminates inside a column, let's split that column
                            workingColumn = lastOverlappingColumn
                                .Split(lastOverlappingColumn.Size - overlappingColumns.OffsetRight);


                            columns.Insert(columns.IndexOf(lastOverlappingColumn) + 1, workingColumn);
                        }
                        else if (overlappingColumns.OffsetRight >= -XThreshold && overlappingColumns.OffsetRight <= XThreshold)
                        {
                            // if control terminates on the border of a column, there's nothing to do
                        }
                        else
                        { // if control terminates past the last overlapping column, we need to create a new column 
                            if (lastOverlappingColumn.NextEntity is null)
                            {
                                workingColumn = new ColumnDefinition
                                {
                                    Size = -overlappingColumns.OffsetRight, // offsets are positive when inside the column
                                                                            // and negative when outside
                                    PreviousEntity = lastOverlappingColumn,
                                };
                                ++additionalSpan;
                                lastOverlappingColumn.NextEntity = workingColumn;
                            }
                            else
                            {
                                throw new InvalidOperationException("Overlap didn't include next column even though offset indicates " +
                                    "it should. This should not happen.");
                            }
                            columns.Insert(columns.IndexOf(lastOverlappingColumn) + 1, workingColumn);
                        }


                        if (overlappingColumns.OffsetLeft > XThreshold)
                        {
                            workingColumn = firstOverlappingColumn
                                .Split(overlappingColumns.OffsetLeft);
                            workingColumn.Objects.Add(control);
                            control.OwningColumn = workingColumn;
                            columns.Insert(columns.IndexOf(firstOverlappingColumn) + 1, workingColumn);
                        }
                        else if ((overlappingColumns.OffsetLeft >= -XThreshold && overlappingColumns.OffsetLeft <= XThreshold))
                        {
                            firstOverlappingColumn.Objects.Add(control);
                            control.OwningColumn = firstOverlappingColumn;
                        }
                        else
                        {

                            /// There's no way to run into this condition because columns are created left-to-right and control
                            /// processing is also done left-to-right, so it would have to be the first column 
                            /// (which is managed on the first branch since at that point no columns have been created) 
                            throw new InvalidOperationException("There cannot be an uncontained left-overlapping column");


                            //if (firstOverlappingColumn.PreviousEntity is null)
                            //{
                            //    throw new InvalidOperationException("There cannot be an left-overlapping column without a previous column");
                            //}
                            //else
                            //{
                            //    throw new InvalidOperationException("There cannot be an left-overlapping column without a previous column");
                            //    workingColumn = firstOverlappingColumn
                            //        .PreviousEntity
                            //        .Split(firstOverlappingColumn.PreviousEntity.Size - overlappingColumns.OffsetLeft);

                            //    workingColumn.Objects.Add(control);
                            //    control.OwningColumn = workingColumn;
                            //    columns.Insert(columns.IndexOf(firstOverlappingColumn), workingColumn);
                            //}
                        }

                        control.ColumnSpan = overlappingColumns.Columns.Count + additionalSpan;
                    }

                }
            }

            return columns;
        }


        /// <summary>
        /// Obtains the columns that a <see cref="VirtualCell"/> spans over
        /// </summary>
        /// <param name="startColumn">The first column in the chain to check</param>
        /// <param name="cell"></param>
        /// <returns>a <see cref="ColumnOverlap"/> defining the overlap</returns>
        private static ColumnOverlap GetOverlappingColumns(ColumnDefinition startColumn, VirtualCell cell)
        {
            var overlapChain = new List<ColumnDefinition>();
            int cellLeftBound = cell.X;
            int cellRightBound = cellLeftBound + cell.Width;

            int leftOffset = 0; // positive when cell starts inside column
            int? rightOffset = null;
            bool chainBegun = false;
            bool chainComplete = false;

            ColumnDefinition? previousColumn = null;
            ColumnDefinition? currentColumn = startColumn;
            while (currentColumn is not null && !chainComplete)
            {
                if (!chainBegun)
                {
                    if (cellRightBound < currentColumn.Bound
                        || (cellLeftBound >= currentColumn.Offset
                            && cellLeftBound < currentColumn.Bound))
                    {
                        leftOffset = cellLeftBound - currentColumn.Offset;
                        chainBegun = true;
                    }
                }

                if (chainBegun)
                {
                    if (cellRightBound <= currentColumn.Bound)
                    {
                        rightOffset = currentColumn.Bound - cellRightBound;
                        chainComplete = true;
                    }
                    overlapChain.Add(currentColumn);
                }

                previousColumn = currentColumn;
                currentColumn = currentColumn.NextEntity;
            }
            rightOffset ??= previousColumn?.Bound - cellRightBound;

            var overlap = new ColumnOverlap(overlapChain, leftOffset, rightOffset ?? -1);

            return overlap;
        }


        private static FloatingCellOffset? MapFloatingControl(IList<ColumnDefinition> columns, IList<RowDefinition> rows, VirtualCell cell)
        {
            int distanceToStartColumn = int.MaxValue;
            int currentOffset;
            ColumnDefinition? startColumn = null;
            ColumnDefinition? endColumn;

            /// get left-bounding column
            foreach (var column in columns)
            {
                currentOffset = cell.X - column.Offset;
                if (currentOffset < distanceToStartColumn && currentOffset >= 0)
                {
                    distanceToStartColumn = currentOffset;
                    startColumn = column;
                }
            }

            if (startColumn is null)
            {
                return null;
            }

            endColumn = startColumn;
            short colSpan = 1;
            int distanceToEndColumn = int.MinValue;

            while (endColumn is not null)
            {
                if ((endColumn.Bound - (cell.RightBound)) >= 0)
                {
                    distanceToEndColumn = endColumn.Bound - cell.RightBound;
                    break;
                }

                endColumn = endColumn.NextEntity;
                ++colSpan;
            }


            if (endColumn is null)
            {
                throw new InvalidOperationException("Could not find rightmost column buffer for floating controls." +
                    " This should not happen");
            }



            ///// Get Y offsets
            //if (cell.OwningRow is null)
            //{
            //    throw new InvalidOperationException("Cell must have an owning row");
            //}

            int distanceToStartRow = int.MaxValue;
            RowDefinition? startRow = null;
            int distanceToCurrentRow;
            foreach (var row in rows)
            {
                distanceToCurrentRow = cell.Object.Y - row.Offset;
                if (distanceToCurrentRow < distanceToStartRow && distanceToCurrentRow >= 0)
                {
                    distanceToStartRow = distanceToCurrentRow;
                    startRow = row;
                }
            }

            if (startRow is null)
                throw new InvalidOperationException("Could not find start row");

            int distanceToEndRow = int.MaxValue;
            RowDefinition? endRow = startRow;
            short rowSpan = 1;
            while (endRow is not null)
            {
                if ((endRow.Bound - cell.LowerBound) >= 0)
                {
                    distanceToEndRow = endRow.Bound - cell.LowerBound;
                    break;
                }

                ++rowSpan;
                endRow = endRow.NextEntity;
            }

            if (endRow is null)
            {
                throw new InvalidOperationException("Could not find rightmost column buffer for floating controls." +
                    " This should not happen");
            }

            return new FloatingCellOffset(startColumn, startRow)
            {
                StartOffset = new Point(distanceToStartColumn, distanceToStartRow),
                EndOffset = new Point(distanceToEndColumn, distanceToEndRow),
                ColSpan = colSpan,
                RowSpan = rowSpan,
            };
        }

    }
}
