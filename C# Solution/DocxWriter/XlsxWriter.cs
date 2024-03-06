using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Appeon.DotnetDemo.DocumentWriter
{
    public class XlsxWriter
    {

        public static int GetStartCell(in XSSFWorkbook workbook, out CellAddress? address, out string? error)
        {
            address = null;
            error = null;

            try
            {
                var sheet = workbook.GetSheetAt(0);

                for (int i = 0; i < 10; ++i)
                {
                    var currentRow = sheet.GetRow(i);
                    if (currentRow is null)
                        continue;
                    for (int j = 0; j < 10; ++j)
                    {
                        var cell = currentRow.GetCell(j);
                        if (cell is null)
                            continue;
                        if (cell.CellType != CellType.Formula && cell.StringCellValue == "###start###")
                        {
                            address = cell.Address;
                            return 1;
                        }
                    }

                }
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }


            address = new CellAddress(0, 0);
            return 1;
        }

        public static int OpenTemplate(string templatePath, out IWorkbook? workbook, out string? error)
        {
            error = null;
            workbook = null;
            try
            {
                using var stream = File.OpenRead(templatePath);

                workbook = new XSSFWorkbook(stream);
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;

                return -1;
            }
        }

        public static int CreateWorkbook(out IWorkbook? workbook, out string? error)
        {
            workbook = null;
            error = null;

            try
            {
                workbook = new XSSFWorkbook();
                workbook.CreateSheet();
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }

        public static int WriteStringArrayToWorkbook(
            ref IWorkbook workbook,
            string[] data,
            string separator,
            out string? error)
        {
            return WriteStringArrayToWorkbookInternal(workbook, data, separator, new CellAddress(0, 0), out error);
        }

        public static int WriteStringArrayToWorkbook(
            ref IWorkbook workbook,
            string[] data,
            string separator,
            CellAddress cellAddress,
            out string? error)
        {
            return WriteStringArrayToWorkbookInternal(workbook, data, separator, cellAddress, out error);
        }

        private static int WriteStringArrayToWorkbookInternal(
            IWorkbook workbook,
            string[] data,
            string separator,
            CellAddress startAddress,
            out string? error)
        {
            error = null;

            try
            {
                var sheet = workbook.GetSheetAt(0);
                IRow? row, previousRow = null;
                int realRow, realColumn, templateDataRows = 0;
                string[] columns;
                ICell? cell, cellAbove, previousCell;
                bool wroteHeader = false, freshRow, finishedRecordingTemplateRange = false;

                for (int i = 0; i < data.Length; ++i, wroteHeader = i > 0)
                {
                    realRow = i + startAddress.Row;
                    if (wroteHeader && templateDataRows > 0)
                        previousRow = sheet.GetRow(realRow - templateDataRows);
                    row = sheet.GetRow(realRow);
                    freshRow = row is null && wroteHeader;
                    if (row is not null && !finishedRecordingTemplateRange)
                    {
                        if (wroteHeader)
                            ++templateDataRows;
                    }
                    else
                    {
                        row = CreateRowCopyStyle(ref sheet, realRow - 1, realRow);
                        finishedRecordingTemplateRange = true;
                        //row = sheet.CreateRow(realRow);
                    }

                    columns = data[i].Split(separator);

                    previousCell = null;
                    for (int j = 0; j < columns.Length; j++)
                    {
                        realColumn = j + startAddress.Column;
                        cell = row.GetCell(realColumn);
                        if (cell is null)
                        {
                            cell = row.CreateCell(realColumn);
                            if (previousCell is not null)
                            {
                                cell.CellStyle = previousCell.CellStyle;
                            }
                        }
                        cell.SetCellValue(columns[j]);
                        if (freshRow && previousRow is not null)
                        {
                            cellAbove = previousRow.GetCell(realColumn);
                            cell.CellStyle = cellAbove.CellStyle;
                        }
                        previousCell = cell;
                    }
                }
                return 1;
            }
            catch (Exception e)
            {
                error = e.Message;
                return -1;
            }
        }

        static IRow CreateRowCopyStyle(ref ISheet sheet, int rowToCopy, int newRowIndex)
        {
            var sourceRow = sheet.GetRow(rowToCopy);

            var newRow = sheet.CreateRow(newRowIndex);

            if (sourceRow is not null)
                for (int i = sourceRow.FirstCellNum; i < sourceRow.LastCellNum; ++i)
                {
                    var newCell = newRow.CreateCell(i);
                    ICell? sourceCell = null;
                    try
                    {
                        sourceCell = sourceRow.GetCell(i);
                    }
                    catch
                    { }

                    if (sourceCell is null)
                        continue;
                    newCell.CellStyle = sourceCell.CellStyle;
                }



            return newRow;
        }

        public static int SaveDocument(in IWorkbook workbook, string fileName, out string? error)
        {
            error = null;

            try
            {
                using FileStream stream = File.Create(fileName);
                workbook.Write(stream, false);
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
