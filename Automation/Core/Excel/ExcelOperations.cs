namespace Automation.Core.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Immutable;
    using System.Drawing;
    using System.Linq;
    using Argument.Check;
    using Microsoft.Office.Interop.Excel;
    using DataTable = System.Data.DataTable;

    /// <summary>
    /// This class deals with all operations related to Excel which includes but is not limited to:
    /// Creating a new workbook, adding a new sheet, opening an existing workbook, finding the sheet,
    /// performing operations like reading and writing to excel and also protecting and unprotecting sheets.
    /// </summary>
    public class ExcelOperations
    {
        private readonly _Application _excel;
        private Workbook _wb;
        private Worksheet _ws;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelOperations"/> class.
        /// </summary>
        public ExcelOperations()
        {
            _excel = new Application
            {
                DisplayAlerts = false,
            };
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelOperations"/> class.
        /// </summary>
        /// <param name="savePath">New excel object save path.</param>
        public ExcelOperations(string savePath)
        {
            _excel = new Application
            {
                DisplayAlerts = false,
            };
            CreateNewFile();
            SaveAs(savePath);
            Close();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelOperations"/> class.
        /// When file is password protected for read as well as write access.
        /// </summary>
        /// <param name="path">Workbook path.</param>
        /// <param name="sheetNumber">Sheet number.</param>
        /// <param name="openFilePassword">Password to open file.</param>
        /// <param name="writePassword">Password to enable editing.</param>
        /// <param name="updateLinksType">Handle how to update Links of Excel File.</param>
        public ExcelOperations(string path, int sheetNumber, string openFilePassword = "", string writePassword = "", UpdateLinksType updateLinksType = UpdateLinksType.Default)
        {
            _excel = new Application();
            object updateLink = Type.Missing;
            if (updateLinksType != UpdateLinksType.Default)
            {
                updateLink = updateLinksType;
            }

            if (!string.IsNullOrEmpty(openFilePassword) && string.IsNullOrEmpty(writePassword))
            {
                _wb = _excel.Workbooks.Open(path, updateLink, Password: openFilePassword);
            }
            else if (!string.IsNullOrEmpty(openFilePassword) && !string.IsNullOrEmpty(writePassword))
            {
                _wb = _excel.Workbooks.Open(path, updateLink, ReadOnly: false, Password: openFilePassword, WriteResPassword: writePassword);
            }
            else
            {
                _wb = _excel.Workbooks.Open(path, updateLink);
            }

            _ws = _wb.Worksheets[sheetNumber];
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelOperations"/> class.
        /// When file is password protected for read as well as write access.
        /// </summary>
        /// <param name="path">Workbook path.</param>
        /// <param name="sheetName">Sheet name.</param>
        /// <param name="openFilePassword">Password to open file.</param>
        /// <param name="writePassword">Password to enable editing.</param>
        /// <param name="updateLinksType">Handle how to update Links of Excel File.</param>
        public ExcelOperations(string path, string sheetName, string openFilePassword = "", string writePassword = "", UpdateLinksType updateLinksType = UpdateLinksType.Default)
        {
            _excel = new Application();
            object updateLink = Type.Missing;
            if (updateLinksType != UpdateLinksType.Default)
            {
                updateLink = updateLinksType;
            }

            if (!string.IsNullOrEmpty(openFilePassword) && string.IsNullOrEmpty(writePassword))
            {
                _wb = _excel.Workbooks.Open(path, updateLink, Password: openFilePassword);
            }
            else if (!string.IsNullOrEmpty(openFilePassword) && !string.IsNullOrEmpty(writePassword))
            {
                _wb = _excel.Workbooks.Open(path, updateLink, ReadOnly: false, Password: openFilePassword, WriteResPassword: writePassword);
            }
            else
            {
                _wb = _excel.Workbooks.Open(path, updateLink);
            }

            _ws = _wb.Worksheets[sheetName];
        }

        /// <summary>
        /// Get the name of the column in excel from the column number.
        /// </summary>
        /// <param name="columnNumber">The 1-based index of the column.</param>
        /// <returns>The name of the column.</returns>
        public static string GetColumnNamePattern(int columnNumber)
        {
            string columnName = string.Empty;
            while (columnNumber > 0)
            {
                int rem = columnNumber % 26;
                if (rem == 0)
                {
                    columnName = "Z" + columnName;
                    columnNumber = (columnNumber / 26) - 1;
                }
                else
                {
                    columnName = Convert.ToChar((rem - 1) + 'A') + columnName;
                    columnNumber = columnNumber / 26;
                }
            }

            return columnName;
        }

        /// <summary>
        /// Open the workbook with the specified path.
        /// </summary>
        /// <param name="path">Path of the workbook.</param>
        /// <param name="updateLinksType">Handle how to update Links of Excel File.</param>
        public void OpenWorkbook(string path, UpdateLinksType updateLinksType = UpdateLinksType.Default)
        {
            object updateLink = Type.Missing;
            if (updateLinksType != UpdateLinksType.Default)
            {
                updateLink = updateLinksType;
            }

            _wb = _excel.Workbooks.Open(path, updateLink);
        }

        /// <summary>
        /// Open the protected workbook with the specified path.
        /// </summary>
        /// <param name="path">Path of the workbook.</param>
        /// <param name="password">Password of Workbook.</param>
        public void OpenProtectedWorkbook(string path, string password)
        {
            _wb = _excel.Workbooks.Open(path, Password: password);
        }

        /// <summary>
        /// Save workbook.
        /// </summary>
        public void Save()
        {
            _wb.Save();
        }

        /// <summary>
        /// Save as new workbook.
        /// </summary>
        /// <param name="path">Path to save workbook.</param>
        public void SaveAs(string path)
        {
            _wb.SaveAs(path);
        }

        /// <summary>
        /// Saves a File as Excel.
        /// </summary>
        /// <param name="path">Path of Excel to be saved.</param>
        public void SaveFileAsExcel(string path)
        {
            _wb.SaveAs(path, XlFileFormat.xlOpenXMLWorkbook);
        }

        /// <summary>
        /// Exports Workbook as PDF.
        /// </summary>
        /// <param name="savePath">Path to save exported PDF file.</param>
        public void ExportAsPDF(string savePath)
        {
            _wb.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, savePath);
        }

        /// <summary>
        /// Closes the workbook and exits excel application.
        /// </summary>
        public void Close()
        {
            _wb?.Close();
            _excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excel);
        }

        /// <summary>
        /// Closes the workbook, without saving if made any changes and exits excel application.
        /// </summary>
        public void CloseWithoutSave()
        {
            object misValue = System.Reflection.Missing.Value;
            _wb.Close(false, misValue, misValue);
            _excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excel);
        }

        /// <summary>
        /// Create a new workbook.
        /// </summary>
        public void CreateNewFile()
        {
            _wb = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _ws = _wb.Worksheets[1];
        }

        /// <summary>
        /// Create new sheet after active sheet.
        /// </summary>
        /// <param name="sheetName">Name of new WorkSheet.</param>
        public void CreateNewSheet(string sheetName = null)
        {
            Worksheet newWorkSheet;
            if (_ws == null)
            {
                newWorkSheet = _wb.Worksheets.Add();
            }
            else
            {
                newWorkSheet = _wb.Worksheets.Add(After: _ws) as Worksheet;
            }

            if (!string.IsNullOrEmpty(sheetName))
            {
                newWorkSheet.Name = sheetName;
            }
        }

        /// <summary>
        /// Activates sheet by sheet index.
        /// </summary>
        /// <param name="sheetIndex">Sheet index.</param>
        public void ActivateSheet(int sheetIndex)
        {
            _ws = _wb.Worksheets[sheetIndex];
        }

        /// <summary>
        /// Activates sheet by sheet index.
        /// </summary>
        /// <param name="sheetName">Sheet Name.</param>
        public void ActivateSheet(string sheetName)
        {
            _ws = _wb.Worksheets[sheetName];
        }

        /// <summary>
        /// Select and Activate specified worksheet number.
        /// </summary>
        /// <param name="sheetNum">Sheet to select.</param>
        public void SelectWorksheet(int sheetNum)
        {
            _ws = _wb.Worksheets[sheetNum];
        }

        /// <summary>
        /// Protect sheet without password.
        /// </summary>
        public void ProtectSheet()
        {
            _ws.Protect();
        }

        /// <summary>
        /// Protect sheet with password provided.
        /// </summary>
        /// <param name="password">Password to protect sheet.</param>
        public void ProtectSheet(string password)
        {
            _ws.Protect(password);
        }

        /// <summary>
        /// Unprotect sheet without password.
        /// </summary>
        public void UnprotectSheet()
        {
            _ws.Unprotect();
        }

        /// <summary>
        /// Unprotect sheet with password provided.
        /// </summary>
        /// <param name="password">Password to unprotect sheet.</param>
        public void UnprotectSheet(string password)
        {
            _ws.Unprotect(password);
        }

        /// <summary>
        /// Check if sheet with specified name exists.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>True if sheet exists else false.</returns>
        public bool IsSheetPresent(string sheetName)
        {
            try
            {
                _ = _wb.Worksheets[sheetName];

                return true;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return false;
            }
        }

        /// <summary>
        /// Get the name of current active sheet.
        /// </summary>
        /// <returns>Current active sheet name.</returns>
        public string GetActiveSheetName()
        {
            return _ws.Name;
        }

        /// <summary>
        /// Get sheet index by name.
        /// </summary>
        /// <param name="sheetName">Sheet name.</param>
        /// <returns>Sheet index.</returns>
        public int GetSheetIndexByName(string sheetName)
        {
            int totalSheets = _wb.Worksheets.Count;
            for (int sheetNumber = 1; sheetNumber <= totalSheets; sheetNumber++)
            {
                Worksheet workSheet = _wb.Worksheets[sheetNumber];
                if (workSheet.Name.Equals(sheetName))
                {
                    return sheetNumber;
                }
            }

            return -1;
        }

        /// <summary>
        /// Get number of sheets in excel.
        /// </summary>
        /// <returns>Sheets count in excel file.</returns>
        public int GetSheetCount()
        {
            return _wb.Worksheets.Count;
        }

        /// <summary>
        /// Create and return a list of sheet names.
        /// </summary>
        /// <returns>Immutable list of sheet names.</returns>
        public ImmutableList<string> GetSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (Worksheet worksheet in _wb.Worksheets)
            {
                sheetNames.Add(worksheet.Name.ToString());
            }

            return sheetNames.ToImmutableList();
        }

        /// <summary>
        /// Create and return a list of visible sheet names.
        /// </summary>
        /// <returns>Immutable list of visible sheet names.</returns>
        public ImmutableList<string> GetVisibleSheetNames()
        {
            List<string> sheetNames = new List<string>();
            foreach (Worksheet worksheet in _wb.Worksheets)
            {
                if (worksheet.Visible == XlSheetVisibility.xlSheetVisible)
                {
                    sheetNames.Add(worksheet.Name.ToString());
                }
            }

            return sheetNames.ToImmutableList();
        }

        /// <summary>
        /// Create and return a dictionary of sheet names and index.
        /// </summary>
        /// <returns>Immutable dictionary of Sheet names and Index.</returns>
        public ImmutableDictionary<string, int> GetMapOfSheetNameAndIndex()
        {
            int sheetIndex = 0;
            Dictionary<string, int> sheetNameAndIndexMap = new Dictionary<string, int>();
            foreach (Worksheet worksheet in _wb.Worksheets)
            {
                sheetNameAndIndexMap.Add(worksheet.Name.ToString(), sheetIndex++);
            }

            return sheetNameAndIndexMap.ToImmutableDictionary();
        }

        /// <summary>
        /// Rename the sheet number specified with the given name.
        /// </summary>
        /// <param name="sheetNumber">Position of sheet in workbook.</param>
        /// <param name="newName">Name with which sheet name needs to be replaced.</param>
        public void RenameSheet(int sheetNumber, string newName)
        {
            Worksheet worksheet = _wb.Worksheets[sheetNumber];
            worksheet.Name = newName;
        }

        /// <summary>
        /// Rename the sheet number specified with the given name.
        /// </summary>
        /// <param name="newSheetName">Name with which sheet name needs to be replaced.</param>
        public void RenameActiveSheet(string newSheetName)
        {
            Worksheet worksheet = _wb.ActiveSheet;
            worksheet.Name = newSheetName;
        }

        /// <summary>
        /// Delete specified worksheet number.
        /// </summary>
        /// <param name="sheetNum">Position of sheet to delete.</param>
        public void DeleteWorksheet(int sheetNum)
        {
            _excel.DisplayAlerts = false;
            _wb.Worksheets[sheetNum].Delete();
            _excel.DisplayAlerts = true;
        }

        /// <summary>
        /// Delete specified worksheet.
        /// </summary>
        /// <param name="sheetName">Name of sheet to delete.</param>
        public void DeleteWorksheet(string sheetName)
        {
            _excel.DisplayAlerts = false;
            _wb.Worksheets[sheetName].Delete();
            _excel.DisplayAlerts = true;
        }

        /// <summary>
        /// Method to get number of rows present in an excel sheet.
        /// </summary>
        /// <returns>An integer number representing no of rows in a sheet.</returns>
        public int GetNumberOfRows()
        {
            return _ws.UsedRange.Rows.Count;
        }

        /// <summary>
        /// Method to get number of col present in an excel sheet.
        /// </summary>
        /// <returns>An integer number representing no of cols in a sheet.</returns>
        public int GetNumberOfCols()
        {
            return _ws.UsedRange.Columns.Count;
        }

        /// <summary>
        /// Insert new rows before a given row.
        /// </summary>
        /// <param name="startPos">Row from where to start insertion.</param>
        /// <param name="noOfRows">Number of Rows to insert.</param>
        public void InsertNewRowsAbove(int startPos, int noOfRows)
        {
            int rowsAdded = 0;
            while (rowsAdded < noOfRows)
            {
                _ws.Rows[startPos].Insert(XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                rowsAdded++;
            }
        }

        /// <summary>
        /// Insert a new column after the given column number and shifts columns right.
        /// </summary>
        /// <param name="col">Column number to add after.</param>
        public void InsertColAndShiftColsRight(int col)
        {
            CheckCellIndices(1, col);
            Range range = _ws.Range[_ws.Cells[1, col], _ws.Cells[GetNumberOfRows(), col]];
            range.Insert(XlInsertShiftDirection.xlShiftToRight);
            _ws.Range[_ws.Cells[1, col], _ws.Cells[GetNumberOfRows(), col]].Value2 = string.Empty;
        }

        /// <summary>
        /// Delete the given row in active worksheet.
        /// </summary>
        /// <param name="row">Row to be deleted.</param>
        public void DeleteRow(int row)
        {
            _ws.Rows[row + 1].Delete();
        }

        /// <summary>
        /// Delete entire row with cell.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Column number.</param>
        public void DeleteRowWithCell(int row, int col)
        {
            Range rng = (Range)_ws.Cells[row, col];
            rng.EntireRow.Delete(Type.Missing);
        }

        /// <summary>
        /// Delete specified row count from top and shift rows up.
        /// </summary>
        /// <param name="rowCount">Row count to delete.</param>
        public void DeleteRowsFromTop(int rowCount)
        {
            int row = 1;
            foreach (Range range in _ws.UsedRange)
            {
                range.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                row++;
                if (row >= rowCount)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Deletes the column with given column number and shifts columns left.
        /// </summary>
        /// <param name="col">Column number to delete.</param>
        public void DeleteColAndShiftColsLeft(int col)
        {
            if (col >= 1)
            {
                Range range = _ws.Range[_ws.Cells[1, col], _ws.Cells[GetNumberOfRows(), col]];
                range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
            }
        }

        /// <summary>
        /// Deletes all the rows in the specified range and shifts rows up.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="noOfRows">Number of Rows to delete.</param>
        public void DeleteRows(int startRow, int noOfRows)
        {
            int endrow = startRow + noOfRows - 1;
            Range range = _ws.Rows[startRow + ":" + endrow];
            range.Delete(XlDeleteShiftDirection.xlShiftUp);
        }

        /// <summary>
        /// Clear all the data and formatting of the excel file active sheet.
        /// </summary>
        /// <param name="clearHeaders">True to clear the header row as well, false otherwise.</param>
        public void ClearActiveSheet(bool clearHeaders = false)
        {
            Range c1 = clearHeaders ? _ws.Cells[1, 1] : _ws.Cells[2, 1];
            Range c2 = _ws.Cells[_ws.UsedRange.Rows.Count, _ws.UsedRange.Columns.Count];
            Range range = _ws.get_Range(c1, c2);
            range.Cells.Clear();
        }

        /// <summary>
        /// Clear all the data and formatting of cells in range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        public void ClearRange(int startRow, int startCol, int endRow, int endCol)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Cells.Clear();
        }

        /// <summary>
        /// Clear only the data of cells in range. Keeps the formatting as it is.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        public void ClearRangeContents(int startRow, int startCol, int endRow, int endCol)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Cells.ClearContents();
        }

        /// <summary>
        /// Set Row Height.
        /// </summary>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="height">Height of the row.</param>
        public void SetRowHeight(int rowIndex, double height)
        {
            Range range = _ws.Rows[rowIndex];
            range.RowHeight = height;
        }

        /// <summary>
        /// Set the width of the given excel column.
        /// </summary>
        /// <param name="columnNumber">Column number.</param>
        /// <param name="width">Width to set.</param>
        public void SetColumnWidth(int columnNumber, double width)
        {
            Range range = _ws.Columns[columnNumber];
            range.ColumnWidth = width;
        }

        /// <summary>
        /// Set the width of excel columns.
        /// </summary>
        /// <param name="width">Width to set.</param>
        public void SetColumnsWidth(double width)
        {
            Range range = _ws.Columns;
            range.ColumnWidth = width;
        }

        /// <summary>
        /// Auto Fit the Used Range of Col and rows.
        /// </summary>
        public void AutoFit()
        {
            _ws.UsedRange.Columns.AutoFit();
            _ws.UsedRange.Rows.AutoFit();
        }

        /// <summary>
        /// Apply Wrap Text to the used range of the worksheet.
        /// </summary>
        public void WrapText()
        {
            _ws.UsedRange.WrapText = true;
        }

        /// <summary>
        /// Wrap Text of the range.
        /// </summary>
        /// <param name="startRow">Start row of the data source sheet.</param>
        /// <param name="startCol">Start column of the data source sheet.</param>
        /// <param name="endRow">End row of the data source sheet.</param>
        /// <param name="endCol">End column of the data source sheet.</param>
        /// <param name="wrapStatus">Boolean to specify whether to wrap or unwrap.</param>
        public void SetRangeWrapText(int startRow, int startCol, int endRow, int endCol, bool wrapStatus = true)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Cells.WrapText = wrapStatus;
        }

        /// <summary>
        /// Sort used range using a column.
        /// </summary>
        /// <param name="colName">Column name to sort entire data by.</param>
        /// <param name="xlSortOrder">Sorting order.</param>
        /// <param name="matchCase">True if case needs to matched.</param>
        /// <exception cref="Exception">Throw exception if specified column is not present.</exception>
        public void SortUsingColumn(string colName, XlSortOrder xlSortOrder, bool matchCase = false)
        {
            int colIndex = GetColumnNumber(colName);
            if (colIndex < 1)
            {
                throw new Exception($"No Column with column name {colName}");
            }

            _ws.UsedRange.Select();
            _ws.Sort.SortFields.Clear();
            _ws.Sort.SortFields.Add(
                _ws.UsedRange.Columns[colIndex],
                XlSortOn.xlSortOnValues,
                xlSortOrder,
                Type.Missing,
                XlSortDataOption.xlSortNormal);
            Sort sort = _ws.Sort;
            sort.SetRange(_ws.UsedRange);
            sort.Header = XlYesNoGuess.xlYes;
            sort.MatchCase = matchCase;
            sort.Orientation = XlSortOrientation.xlSortColumns;
            sort.SortMethod = XlSortMethod.xlPinYin;
            sort.Apply();
        }

        /// <summary>
        /// Turn all filters of Active sheet Off.
        /// </summary>
        public void RemoveFiltersOfActiveSheet()
        {
            if (_ws.AutoFilter != null)
            {
                _ws.AutoFilterMode = false;
            }
        }

        /// <summary>
        /// Get the column index by using the cell value and the row number(1-based).
        /// </summary>
        /// <param name="cellValue">Cell value.</param>
        /// <param name="rowNumber">The index(1-based) of the row, default value is 1.</param>
        /// <returns>Position of the column.</returns>
        public int GetColumnNumber(string cellValue, int rowNumber = 1)
        {

            int totalColumns = _ws.Cells[rowNumber, _ws.Columns.Count].End(XlDirection.xlToLeft).Column;
            for (int colNumber = 1; colNumber <= totalColumns; colNumber++)
            {
                if (ReadCell(rowNumber, colNumber).Equals(cellValue))
                {
                    return colNumber;
                }
            }

            return -1;
        }

        /// <summary>
        /// Get the list of indexes of columns having the cell value in given row number.
        /// </summary>
        /// <param name="cellValue">Name of the column.</param>
        /// <param name="rowNumber">The index(1-based) of the row, default value is 1.</param>
        /// <returns>List of index of the column.</returns>
        public List<int> GetColumnNumbers(string cellValue, int rowNumber = 1)
        {
            List<int> columnNumberList = new List<int>();
            int totalColumns = _ws.Cells[rowNumber, _ws.Columns.Count].End(XlDirection.xlToLeft).Column;
            for (int colNumber = 1; colNumber <= totalColumns; colNumber++)
            {
                if (ReadCell(rowNumber, colNumber).Equals(cellValue))
                {
                    columnNumberList.Add(colNumber);
                }
            }

            return columnNumberList;
        }

        /// <summary>
        /// Get the row index by using the cell value and the column index(1-based).
        /// </summary>
        /// <param name="cellValue">Name of the column.</param>
        /// <param name="columnNumber">The index(1-based) of the column, default value is 1.</param>
        /// <returns>Position of the column.</returns>
        public int GetRowNumber(string cellValue, int columnNumber = 1)
        {
            int totalRows = _ws.Cells[_ws.Rows.Count, columnNumber].End(XlDirection.xlUp).Row;
            for (int rowNumber = 1; rowNumber <= totalRows; rowNumber++)
            {
                if (ReadCell(rowNumber, columnNumber).Equals(cellValue))
                {
                    return rowNumber;
                }
            }

            return -1;
        }

        /// <summary>
        /// Gets the index (1-based) of the last filled row.
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>

        public int GetLastFilledRow(int columnNumber = 1)
        {
            int totalRows = _ws.Cells[_ws.Rows.Count, columnNumber].End(XlDirection.xlUp).Row;

            return totalRows;
        }

        /// <summary>
        /// Add headers to worksheet.
        /// </summary>
        /// <param name="headers">Headers.</param>
        /// <param name="headerRowIndex">The index(1-based) of the header row, default value is 1.</param>
        public void AddHeaders(string[,] headers, int headerRowIndex = 1)
        {
            Throw.IfNull(() => headers);
            WriteRangeObject(headerRowIndex, 1, headerRowIndex, headers.Length, headers);
            SetRangeBackgroundColor(headerRowIndex, 1, headerRowIndex, headers.Length, Color.AliceBlue);
            AutoFit();
        }

        /// <summary>
        /// Add headers to worksheet.
        /// </summary>
        /// <param name="headers">Headers.</param>
        /// <param name="headerRowIndex">The index(1-based) of the header row, default value is 1.</param>
        public void AddHeaders(List<string> headers, int headerRowIndex = 1)
        {
            Throw.IfNull(() => headers);
            string[,] headerWriter = new string[1, headers.Count];
            int colCount = 0;
            foreach (string item in headers)
            {
                headerWriter[0, colCount++] = item;
            }

            WriteRangeObject(headerRowIndex, 1, headerRowIndex, headers.Count, headerWriter);
            SetRangeBackgroundColor(headerRowIndex, 1, headerRowIndex, headers.Count, Color.AliceBlue);
            AutoFit();
        }

        /// <summary>
        /// Fetch the column names as a List of strings.
        /// </summary>
        /// <param name="headerRowIndex">The index(1-based) of the header row.</param>
        /// <returns>List of string.</returns>
        public List<string> GetHeaders(int headerRowIndex = 1)
        {
            List<string> headers = new List<string>();
            int totalColumns = _ws.Cells[headerRowIndex, _ws.Columns.Count].End(XlDirection.xlToLeft).Column;
            for (int colNumber = 1; colNumber <= totalColumns; colNumber++)
            {
                headers.Add(ReadCell(headerRowIndex, colNumber));
            }

            return headers;
        }

        /// <summary>
        /// Fetch the column names as a Dictionary of header and index.
        /// </summary>
        /// <param name="headerRowIndex">The index(1-based) of the header row.</param>
        /// <returns>Dictionary of headers and index.</returns>
        public Dictionary<string, int> GetMapOfHeaderAndIndex(int headerRowIndex = 1)
        {
            Dictionary<string, int> headersAndIndexMap = new Dictionary<string, int>();
            int totalColumns = _ws.Cells[headerRowIndex, _ws.Columns.Count].End(XlDirection.xlToLeft).Column;
            for (int colNumber = 1; colNumber <= totalColumns; colNumber++)
            {
                string key = ReadCell(headerRowIndex, colNumber);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    headersAndIndexMap[key] = colNumber;
                }
            }

            return headersAndIndexMap;
        }

        /// <summary>
        /// Read a particular cell in a row and column.
        /// </summary>
        /// <param name="row">Cell Row.</param>
        /// <param name="col">Cell Column.</param>
        /// <returns>String value of cell content.</returns>
        public string ReadCell(int row, int col)
        {
            CheckCellIndices(row, col);

            return _ws.Cells[row, col].Value2 != null ? (string)_ws.Cells[row, col].Text.ToString() : string.Empty;
        }

        /// <summary>
        /// Read hyperlink from a particular cell in a row and column.
        /// </summary>
        /// <param name="row">Cell Row.</param>
        /// <param name="col">Cell Column.</param>
        /// <returns>Hyperlink value of cell content.</returns>
        public string ReadHyperlinkFromCell(int row, int col)
        {
            CheckCellIndices(row, col);

            return string.IsNullOrWhiteSpace(_ws.Cells[row, col].Value2) ? string.Empty : ((Range)_ws.Cells[row, col]).Cells.Hyperlinks[1].Address;
        }

        /// <summary>
        /// Read formula from a particular cell in a row and column.
        /// </summary>
        /// <param name="row">Cell Row.</param>
        /// <param name="col">Cell Column.</param>
        /// <returns>Formula value of cell content.</returns>
        public string ReadFormulaFromCell(int row, int col)
        {
            CheckCellIndices(row, col);
            string formulaInCell = Convert.ToString(_ws.Cells[row, col].Formula);

            return string.IsNullOrWhiteSpace(formulaInCell) ? string.Empty : formulaInCell;
        }

        /// <summary>
        /// Reads column values and returns as array.
        /// </summary>
        /// <param name="colNum">Column position starting with 1.</param>
        /// <returns>String array of column values.</returns>
        public string[] ReadColumn(int colNum)
        {
            CheckCellIndices(1, colNum);
            Range range = _ws.UsedRange.Columns[colNum];
            if (range.Value == null)
            {
                return new string[0];
            }

            if (range.Cells.Count == 1)
            {
                return new string[] { Convert.ToString(range.Value) };
            }

            Array columnValues = (Array)range.Cells.Value;
            string[] columnValuesArray = columnValues.OfType<object>().Select(o => o.ToString()).ToArray();

            return columnValuesArray;
        }

        /// <summary>
        /// Get the number of rows in the column having index <paramref name="colNum"/>.
        /// </summary>
        /// <param name="colNum">1-based index of the column.</param>
        /// <returns>Count of the rows.</returns>
        public int GetNumberOfRowsInColumn(int colNum)
        {
            CheckCellIndices(1, colNum);
            Range range = _ws.UsedRange.Columns[colNum];

            return range.Cells.Count;
        }

        /// <summary>
        /// Reads column values and returns as array including empty cells.
        /// </summary>
        /// <param name="colNum">Column index or name starting with 0.</param>
        /// <returns>String array of column values.</returns>
        public string[] ReadColumnIncludingEmptyCells(int colNum)
        {
            Range range = _ws.UsedRange.Columns[colNum];
            if (range.Value == null)
            {
                return new string[0];
            }

            if (range.Cells.Count == 1)
            {
                return new string[] { Convert.ToString(range.Value) };
            }

            dynamic values = range.Cells.Value;
            string[] columnData = new string[values.Length];
            for (int index = 1; index <= values.Length; index++)
            {
                if (values.GetValue(index, 1) == null)
                {
                    columnData[index - 1] = string.Empty;
                }
                else
                {
                    columnData[index - 1] = Convert.ToString(values.GetValue(index, 1));
                }
            }

            return columnData;
        }

        /// <summary>
        /// Reads specified range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <returns>2D string matrix of read range.</returns>
        public string[,] ReadRange(int startRow, int startCol, int endRow, int endCol)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            if (range.Cells.Count == 1)
            {
                return new string[1, 1] { { range.Value2 != null ? Convert.ToString(range.Value2) : string.Empty } };
            }

            object[,] holder = range.Value2;
            string[,] stringHolder = new string[endRow - startRow + 1, endCol - startCol + 1];

            for (int i = 1; i <= endRow - startRow + 1; i++)
            {
                for (int j = 1; j <= endCol - startCol + 1; j++)
                {
                    stringHolder[i - 1, j - 1] = holder[i, j] != null ? holder[i, j].ToString() : string.Empty;
                }
            }

            return stringHolder;
        }

        /// <summary>
        /// Reads specified range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <returns>2D object matrix of read range.</returns>
        public object[,] ReadRangeObject(int startRow, int startCol, int endRow, int endCol)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            if (range.Cells.Count == 1)
            {
                return new object[1, 1] { { range.Value2 } };
            }

            object[,] holder = range.Value2;
            object[,] objectHolder = new object[endRow - startRow + 1, endCol - startCol + 1];

            for (int i = 1; i <= endRow - startRow + 1; i++)
            {
                for (int j = 1; j <= endCol - startCol + 1; j++)
                {
                    objectHolder[i - 1, j - 1] = holder[i, j];
                }
            }

            return objectHolder;
        }

        /// <summary>
        /// Write string value to particular cell in row and column.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Col number.</param>
        /// <param name="value">Value to write.</param>
        /// <param name="color">Cell background color. Optional parameter, no background color by default.</param>
        public void WriteStringToCell(int row, int col, string value, Color? color = null)
        {
            CheckCellIndices(row, col);
            _ws.Cells[row, col].Value2 = value;
            if (color != null)
            {
                SetCellBackgroundColor(row, col, color.Value);
            }
        }

        /// <summary>
        /// Write Object value to particular cell in row and column.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Col number.</param>
        /// <param name="value">Value to write.</param>
        /// <param name="color">Cell background color. Optional parameter, no background color by default.</param>
        public void WriteObjectToCell(int row, int col, object value, Color? color = null)
        {
            CheckCellIndices(row, col);
            _ws.Cells[row, col].Value2 = value;
            if (color != null)
            {
                SetCellBackgroundColor(row, col, color.Value);
            }
        }

        /// <summary>
        /// Write formula value to particular cell in row and column.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Col number.</param>
        /// <param name="formula">Formula to write.</param>
        /// <param name="color">Cell background color. Optional parameter, no background color by default.</param>
        public void WriteFormulaToCell(int row, int col, string formula, Color? color = null)
        {
            CheckCellIndices(row, col);
            _ws.Cells[row, col].Formula = formula;
            if (color != null)
            {
                SetCellBackgroundColor(row, col, color.Value);
            }
        }

        /// <summary>
        /// Write 2D object matrix to specified range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="writeMatrix">2D string matrix to be written to sheet.</param>
        /// <param name="autofit">Auto fit.</param>
        public void WriteRangeObject(int startRow, int startCol, int endRow, int endCol, object[,] writeMatrix, bool autofit = true)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Value2 = writeMatrix;
            if (autofit)
            {
                AutoFit();
            }
        }

        /// <summary>
        /// Write 2D string matrix to specified range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="writeMatrix">2D string matrix to be written to sheet.</param>
        public void WriteRangeString(int startRow, int startCol, int endRow, int endCol, string[,] writeMatrix)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Value2 = writeMatrix;
            AutoFit();
        }

        /// <summary>
        /// Write 1D string matrix to specified range after transpose.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="writeMatrix">1D string matrix to be written to sheet.</param>
        public void WriteRangeStringTranspose(int startRow, int startCol, int endRow, int endCol, string[] writeMatrix)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Value = _excel.WorksheetFunction.Transpose(writeMatrix);
        }

        /// <summary>
        /// Write datatable to excel with headers.
        /// </summary>
        /// <param name="dataTable">DataTable to be written to excel.</param>
        public void WriteDataTableToSheet(DataTable dataTable)
        {
            Throw.IfNull(() => dataTable);
            for (int i = 1; i < dataTable.Columns.Count + 1; i++)
            {
                _ws.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
            }

            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int k = 0; k < dataTable.Columns.Count; k++)
                {
                    _ws.Cells[j + 2, k + 1] = dataTable.Rows[j].ItemArray[k].ToString();
                }
            }

            AutoFit();
        }

        /// <summary>
        /// Write datatable to excel with headers and background color header.
        /// </summary>
        /// <param name="dataTable">DataTable to be written to excel.</param>
        /// <param name="color">Color.</param>
        public void WriteDataTableToSheetWithHeaderColor(DataTable dataTable, Color color)
        {
            Throw.IfNull(() => dataTable);
            for (int i = 1; i < dataTable.Columns.Count + 1; i++)
            {
                _ws.Cells[1, i] = dataTable.Columns[i - 1].ColumnName;
                SetRangeBackgroundColor(1, i, 1, i, color);
            }

            for (int j = 0; j < dataTable.Rows.Count; j++)
            {
                for (int k = 0; k < dataTable.Columns.Count; k++)
                {
                    _ws.Cells[j + 2, k + 1] = dataTable.Rows[j].ItemArray[k].ToString();
                }
            }

            AutoFit();
        }

        /// <summary>
        /// Rename the sheet number specified with the given name.
        /// </summary>
        /// <param name="sheetName">Sheet name which need to be copied.</param>
        /// <param name="copiedSheetName">Copied sheetName.</param>
        public void CopySheet(string sheetName, string copiedSheetName)
        {
            if (IsSheetPresent(sheetName))
            {
                Worksheet worksheet = _wb.Worksheets[sheetName];
                worksheet.Copy(After: worksheet);
                RenameActiveSheet(copiedSheetName);
            }
            else
            {
                throw new Exception($"Sheet name: {sheetName} is not present.");
            }
        }

        /// <summary>
        /// Copy all data from one Sheet to another between 2 different files, Including all the formatting and formulas.
        /// </summary>
        /// <param name="path">Path of file other than open file.</param>
        /// <param name="sourceSheet">Copy from sheet name.</param>
        /// <param name="sheetNumber">Sheet number of destination file.</param>
        public void CopyFromAnotherFile(string path, string sourceSheet, int sheetNumber)
        {
            Workbook wb = null;
            try
            {
                wb = _excel.Workbooks.Open(path);
                Worksheet ws = wb.Worksheets[sourceSheet];
                Range copyRange = ws.Rows[1 + ":" + ws.UsedRange.Rows.Count];
                _ws = _wb.Worksheets[sheetNumber];
                Range dest = _ws.Rows[1 + ":" + ws.UsedRange.Rows.Count];
                copyRange.Copy(dest);
                AutoFit();
            }
            finally
            {
                wb?.Close(false);
            }
        }

        /// <summary>
        /// Copy all data from one Sheet to another between 2 different files, Including all the formatting and formulas.
        /// </summary>
        /// <param name="path">Path of file other than open file.</param>
        /// <param name="sourceSheetNum">Copy from sheet number.</param>
        /// <param name="sheetName">Sheet name of destination file.</param>
        public void CopyFromAnotherFile(string path, int sourceSheetNum, string sheetName)
        {
            Workbook wb = null;
            try
            {
                wb = _excel.Workbooks.Open(path);
                Worksheet ws = wb.Worksheets[sourceSheetNum];
                Range copyRange = ws.Rows[1 + ":" + ws.UsedRange.Rows.Count];
                _ws = _wb.Worksheets[sheetName];
                Range dest = _ws.Rows[1 + ":" + ws.UsedRange.Rows.Count];
                copyRange.Copy(dest);
                AutoFit();
            }
            finally
            {
                wb?.Close(false);
            }
        }

        /// <summary>
        /// Copy all data from one Sheet to another between 2 different files, Including all the formatting and formulas.
        /// </summary>
        /// <param name="path">Path of file other than open file.</param>
        /// <param name="destSheetNumber">Copy to sheet name.</param>
        /// <param name="sourceSheetNumber">Sheet number of source file.</param>
        public void CopySheetDataToAnotherFile(string path, int destSheetNumber, int sourceSheetNumber)
        {
            Workbook wb = null;
            try
            {
                wb = _excel.Workbooks.Open(path);
                Worksheet ws = wb.Worksheets[destSheetNumber];
                _ws = _wb.Worksheets[sourceSheetNumber];
                Range copyRange = _ws.Rows[1 + ":" + _ws.UsedRange.Rows.Count];
                Range dest = ws.Rows[1 + ":" + ws.UsedRange.Rows.Count];
                copyRange.Copy(dest);
                wb.Save();
            }
            finally
            {
                wb?.Close(false);
            }
        }

        /// <summary>
        /// Merge cells in Excel sheets in the given range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        public void MergeRangeCells(int startRow, int startCol, int endRow, int endCol)
        {
            _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]].HorizontalAlignment = XlHAlign.xlHAlignCenter;
            _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]].Merge();
        }

        /// <summary>
        /// Merge cells in Excel sheets in the given range and add background color.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="color">Color.</param>
        public void MergeRangeCellsWithColor(int startRow, int startCol, int endRow, int endCol, Color color)
        {
            MergeRangeCells(startRow, startCol, endRow, endCol);
            SetRangeBackgroundColor(startRow, startCol, endRow, endCol, color);
        }

        /// <summary>
        /// Set the Font color of Range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="color">Color.</param>
        public void SetRangeFontColor(int startRow, int startCol, int endRow, int endCol, Color color)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Font.Color = color;
        }

        /// <summary>
        /// Set the Background color of a cell.
        /// </summary>
        /// <param name="row">Cell Row.</param>
        /// <param name="col">Cell Col.</param>
        /// <param name="color">color.</param>
        public void SetCellBackgroundColor(int row, int col, Color color)
        {
            Range range = (Range)_ws.Cells[row, col];
            range.Interior.Color = color;
        }

        /// <summary>
        /// Set the Background color of Range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="color">color.</param>
        public void SetRangeBackgroundColor(int startRow, int startCol, int endRow, int endCol, Color color)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Interior.Color = color;
        }

        /// <summary>
        /// Removes the Background color of Range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        public void RemoveRangeBackgroundColor(int startRow, int startCol, int endRow, int endCol)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
        }

        /// <summary>
        /// Set the Number Format of Column.
        /// </summary>
        /// <param name="colNo">Column no.</param>
        /// <param name="format">Number Format.</param>
        public void SetColumnNumberFormat(int colNo, string format)
        {
            CheckCellIndices(1, colNo);
            _ws.Columns[colNo].NumberFormat = format;
        }

        /// <summary>
        /// Set the Number Format of Range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="format">Number Format.</param>
        public void SetRangeNumberFormat(int startRow, int startCol, int endRow, int endCol, string format)
        {
            _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]].NumberFormat = format;
        }

        /// <summary>
        /// Horizontal Alignment of text.
        /// </summary>
        /// <param name="startRow">Start row of the data source sheet.</param>
        /// <param name="startCol">Start column of the data source sheet.</param>
        /// <param name="endRow">End row of the data source sheet.</param>
        /// <param name="endCol">End column of the data source sheet.</param>
        /// <param name="xlHAlign">Alignment Type, default value is XlHAlign.xlHAlignLeft.</param>
        public void SetRangeHorizontalAlignment(int startRow, int startCol, int endRow, int endCol, XlHAlign xlHAlign = XlHAlign.xlHAlignLeft)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.HorizontalAlignment = xlHAlign;
        }

        /// <summary>
        /// Vertical Alignment of text.
        /// </summary>
        /// <param name="startRow">Start row of the data source sheet.</param>
        /// <param name="startCol">Start column of the data source sheet.</param>
        /// <param name="endRow">End row of the data source sheet.</param>
        /// <param name="endCol">End column of the data source sheet.</param>
        /// <param name="xlVAlign">Alignment Type, default value is XlVAlign.xlVAlignCenter.</param>
        public void SetRangeVerticalAlignment(int startRow, int startCol, int endRow, int endCol, XlVAlign xlVAlign = XlVAlign.xlVAlignCenter)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.VerticalAlignment = xlVAlign;
        }

        /// <summary>
        /// Set Text Style To Bold for given cell.
        /// </summary>
        /// <param name="row">Cell row.</param>
        /// <param name="col">Cell column.</param>
        /// <param name="boldStatus">Boolean to specify whether to set to bold or normal, default value is to set to bold.</param>
        public void SetCellTextStyleToBold(int row, int col, bool boldStatus = true)
        {
            Range range = (Range)_ws.Cells[row, col];
            range.Cells.Font.Bold = boldStatus;
        }

        /// <summary>
        /// Set Text Style To Bold for cells in given range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="boldStatus">Boolean to specify whether to set to bold or normal, default value is to set to bold.</param>
        public void SetRangeTextStyleToBold(int startRow, int startCol, int endRow, int endCol, bool boldStatus = true)
        {
            _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]].Cells.Font.Bold = boldStatus;
        }

        /// <summary>
        /// Set border for cells in given range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="borderStyle">Border style, defaults to XlLineStyle.xlContinuous.</param>
        public void SetCellsBorderInRange(int startRow, int startCol, int endRow, int endCol, XlLineStyle borderStyle = XlLineStyle.xlContinuous)
        {
            _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]].Cells.Borders.LineStyle = borderStyle;
        }

        /// <summary>
        /// Set border around a given range.
        /// </summary>
        /// <param name="startRow">Starting row cell.</param>
        /// <param name="startCol">Starting column cell.</param>
        /// <param name="endRow">Ending row cell.</param>
        /// <param name="endCol">Ending column cell.</param>
        /// <param name="borderStyle">Border style, defaults to XlLineStyle.xlContinuous.</param>
        public void SetBorderAroundRange(int startRow, int startCol, int endRow, int endCol, XlLineStyle borderStyle = XlLineStyle.xlContinuous)
        {
            Range borderizeRange = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            Borders border = borderizeRange.Borders;
            border[XlBordersIndex.xlEdgeBottom].LineStyle = borderStyle;
            border[XlBordersIndex.xlEdgeTop].LineStyle = borderStyle;
            border[XlBordersIndex.xlEdgeLeft].LineStyle = borderStyle;
            border[XlBordersIndex.xlEdgeRight].LineStyle = borderStyle;
        }

        /// <summary>
        /// Refresh Pivot Tables of active worksheet.
        /// </summary>
        public void RefreshPivot()
        {
            dynamic pivotTables = _ws.PivotTables();
            int pivotTablesCount = pivotTables.Count;
            for (int count = 1; count <= pivotTablesCount; count++)
            {
                pivotTables.Item(count).RefreshTable();
            }
        }
        /// <summary>
        /// Creates the empty pivot table.
        /// </summary>
        /// <param name="row">Row of the cell where the first field pivot table is to be placed.</param>
        /// <param name="col">Column of the cell where the first field of the pivot table is to be placed.</param>
        /// <param name="startRow">Start row of the data source sheet.</param>
        /// <param name="startCol">Start column of the data source sheet.</param>
        /// <param name="endRow">End row of the data source sheet.</param>
        /// <param name="endCol">End column of the data source sheet.</param>
        /// <param name="dataSourceSheet">Name or index of the sheet from which data has to be picked.</param>
        /// <param name="pivotTableVersion">Pivot Table Version.</param>
        /// <returns>Instance of the PivotTable.</returns>
        public PivotTable AddEmptyPivotTable(
                        int row,
                        int col,
                        int startRow,
                        int startCol,
                        int endRow,
                        int endCol,
                        object dataSourceSheet,
                        XlPivotTableVersionList pivotTableVersion = XlPivotTableVersionList.xlPivotTableVersion15)
        {
            Worksheet sourceSheet = (Worksheet)_wb.Worksheets[dataSourceSheet];
            Range dataRange = sourceSheet.Range[sourceSheet.Cells[startRow, startCol], sourceSheet.Cells[endRow, endCol]];
            PivotCache pivotCache = _wb.PivotCaches().Create(XlPivotTableSourceType.xlDatabase, dataRange, pivotTableVersion);
            return ((PivotTables)_ws.PivotTables()).Add(pivotCache, _ws.Cells[row, col]);
        }

        /// <summary>
        /// Creates the empty pivot table from the complete range of the specified sheet.
        /// </summary>
        /// <param name="row">Row of the cell where the first field pivot table is to be placed.</param>
        /// <param name="col">Column of the cell where the first field of the pivot table is to be placed.</param>
        /// <param name="dataSourceSheet">Name or index of the sheet from which data has to be picked.</param>
        /// <param name="pivotTableVersion">Pivot Table Version.</param>
        /// <returns>Instance of the PivotTable.</returns>
        public PivotTable AddEmptyPivotTable(
            int row, int col, object dataSourceSheet, XlPivotTableVersionList pivotTableVersion = XlPivotTableVersionList.xlPivotTableVersion15)
        {
            Worksheet sourceSheet = (Worksheet)_wb.Worksheets[dataSourceSheet];
            PivotCache pivotCache = _wb.PivotCaches().Create(
                XlPivotTableSourceType.xlDatabase, sourceSheet.UsedRange, pivotTableVersion);
            return ((PivotTables)_ws.PivotTables()).Add(pivotCache, _ws.Cells[row, col]);
        }

        /// <summary>
        /// Change source sheet for pivot table.
        /// </summary>
        /// <param name="pivotTable">Pivot table to modify.</param>
        /// <param name="dataSourceSheet">Name or index of the sheet from which data has to be picked.</param>
        /// <param name="pivotTableVersion">Pivot Table Version.</param>
        public void ChangeSourceSheetForPivotTable(
                                        PivotTable pivotTable,
                                        object dataSourceSheet,
                                        XlPivotTableVersionList pivotTableVersion = XlPivotTableVersionList.xlPivotTableVersion15)
        {
            Worksheet sourceSheet = (Worksheet)_wb.Worksheets[dataSourceSheet];
            Range range = sourceSheet.UsedRange;
            _wb.Names.Add(Name: "RangeForPivot", RefersTo: range);
            PivotCache pivotCache = _wb.PivotCaches().Create(
                                        XlPivotTableSourceType.xlDatabase, "RangeForPivot", pivotTableVersion);
            pivotTable.ChangePivotCache(pivotCache);
        }

        /// <summary>
        /// Get all pivot tables form worksheet.
        /// </summary>
        /// <returns>List of pivot tables in worksheet.</returns>
        public List<PivotTable> GetAllPivotTables()
        {
            dynamic pivotTables = _ws.PivotTables();
            int pivotTablesCount = pivotTables.Count;
            List<PivotTable> pivotTableList = new List<PivotTable>();
            for (int count = 1; count <= pivotTablesCount; count++)
            {
                pivotTableList.Add(pivotTables.Item(count));
            }

            return pivotTableList;
        }

        /// <summary>
        /// Get the total no. of rows in Pivot table.
        /// </summary>
        /// <param name="tableNameOrIndex">Name or index of the Pivot table in the active sheet.</param>
        /// <param name="includePageFields">True to include Page fields in range, false otherwise.</param>
        /// <returns>Number of rows in pivot table.</returns>
        public int GetNumberOfRowsOfPivotTable(object tableNameOrIndex, bool includePageFields = false)
        {
            Throw.IfNull(() => tableNameOrIndex);
            PivotTable table = _ws.PivotTables().Item(tableNameOrIndex);
            return includePageFields ? table.TableRange2.Rows.Count : table.TableRange1.Rows.Count;
        }

        /// <summary>
        /// Get the total no. of columns in Pivot table.
        /// </summary>
        /// <param name="tableNameOrIndex">Name or index of the Pivot table in the active sheet.</param>
        /// <param name="includePageFields">True to include Page fields in range, false otherwise.</param>
        /// <returns>Number of columns in pivot table.</returns>
        public int GetNumberOfColsOfPivotTable(object tableNameOrIndex, bool includePageFields = false)
        {
            Throw.IfNull(() => tableNameOrIndex);
            PivotTable table = _ws.PivotTables().Item(tableNameOrIndex);
            return includePageFields ? table.TableRange2.Columns.Count : table.TableRange1.Columns.Count;
        }

        /// <summary>
        /// Get the start row index (1-based) of the Pivot table.
        /// </summary>
        /// <param name="tableNameOrIndex">Name or index of the Pivot table in the active sheet.</param>
        /// <param name="includePageFields">True to include Page fields in range, false otherwise.</param>
        /// <returns>Starting row index of pivot table.</returns>
        public int GetStartRowIndexOfPivotTable(object tableNameOrIndex, bool includePageFields = false)
        {
            Throw.IfNull(() => tableNameOrIndex);
            PivotTable table = _ws.PivotTables().Item(tableNameOrIndex);
            return includePageFields ? table.TableRange2.Row : table.TableRange1.Row;
        }

        /// <summary>
        /// Get the start column index (1-based) of the Pivot table.
        /// </summary>
        /// <param name="tableNameOrIndex">Name or index of the Pivot table in the active sheet.</param>
        /// <param name="includePageFields">True to include Page fields in range, false otherwise.</param>
        /// <returns>Starting column index of pivot table.</returns>
        public int GetStartColumnIndexOfPivotTable(object tableNameOrIndex, bool includePageFields = false)
        {
            Throw.IfNull(() => tableNameOrIndex);
            PivotTable table = _ws.PivotTables().Item(tableNameOrIndex);
            return includePageFields ? table.TableRange2.Column : table.TableRange1.Column;
        }

        /// <summary>
        /// Set the font size of the range.
        /// </summary>
        /// <param name="startRow">start row index.</param>
        /// <param name="startCol">start coloumn index.</param>
        /// <param name="endRow">end row index.</param>
        /// <param name="endCol">end coloumn index.</param>
        /// <param name="size">Size of the font.</param>
        public void SetRangeFontSize(int startRow, int startCol, int endRow, int endCol, double size)
        {
            Range range = _ws.Range[_ws.Cells[startRow, startCol], _ws.Cells[endRow, endCol]];
            range.Font.Size = size;
        }

        private static void ModifySubtotals(PivotField field, object[] subtotals)
        {
            for (int index = 1; index < 13; index++)
            {
                field.set_Subtotals(index, subtotals[index - 1]);
            }
        }

        private static void CheckCellIndices(int rowNum, int colNum)
        {
            if (rowNum < 1 || colNum < 1)
            {
                throw new Exception("Row & Column indices should start with 1.");
            }
        }
    }
}
