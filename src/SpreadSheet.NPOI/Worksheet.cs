using System;
using System.Collections.Generic;
using NPOI.SS.Util;
using SpreadSheet.Core;
using NPOIWorksheet = global::NPOI.SS.UserModel.ISheet;

namespace SpreadSheet.NPOI {

    public sealed class Worksheet : IWorksheet {

        #region Internal Properties

        internal NPOIWorksheet WorksheetImpl { get; }

        #endregion Internal Properties

        #region Internal Constructors

        internal Worksheet(NPOIWorksheet worksheet) {
            WorksheetImpl = worksheet;
        }

        #endregion Internal Constructors

        #region Private Methods

        private IEnumerable<IColumn> GetColumns() {
            throw new NotImplementedException();
        }

        private IEnumerable<IRow> GetRows() {
            throw new NotImplementedException();
        }

        #endregion Private Methods

        #region IWorksheet Members

        public string Name {
            get { return WorksheetImpl.SheetName; }
            set { WorksheetImpl.Workbook.SetSheetName(WorksheetImpl.Workbook.GetSheetIndex(WorksheetImpl), value); }
        }

        public IEnumerable<IColumn> Columns => GetColumns();

        public IEnumerable<IRow> Rows => GetRows();

        public IColumn CreateColumn(int insertAt = -1) {
            var totalRows = WorksheetImpl.LastRowNum + 1;

            throw new NotImplementedException();
        }

        public IRow CreateRow(int insertAt = -1) {
            WorksheetImpl.CreateRow(insertAt);

            throw new NotImplementedException();
        }

        public IRange GetRange(int startColumn, int startRow, int endColumn, int endRow) {
            var startReference = new CellReference(startRow, startColumn);
            var endReference = new CellReference(endRow, endColumn);
            var rowsCount = (endReference.Row - startReference.Row + 1);
            var columnsCount = (endReference.Col - startReference.Col + 1);

            var cells = new Cell[rowsCount][];

            for (var rowIndex = startReference.Row; rowIndex < (endReference.Row + 1); rowIndex++) {
                var row = WorksheetImpl.GetRow(rowIndex);
                var arrayRowIndex = rowIndex - startReference.Row;
                cells[arrayRowIndex] = new Cell[columnsCount];
                for (var columnIndex = startReference.Col; columnIndex < (endReference.Col + 1); columnIndex++) {
                    var arrayColIndex = columnIndex - startReference.Col;
                    cells[arrayRowIndex][arrayColIndex] = new Cell(row.GetCell(columnIndex));
                }
            }

            return new Range(cells);
        }

        #endregion IWorksheet Members
    }
}