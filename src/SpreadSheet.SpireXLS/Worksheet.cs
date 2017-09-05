using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using SpreadSheet.Core;
using SpireWorksheet = Spire.Xls.Worksheet;

namespace SpreadSheet.SpireXLS {

    public sealed class Worksheet : IWorksheet {

        #region Private Read-Only Fields

        private readonly SpireWorksheet _worksheet;

        #endregion Private Read-Only Fields

        #region Internal Properties

        internal SpireWorksheet WorksheetImpl {
            get { return _worksheet; }
        }

        #endregion Internal Properties

        #region Internal Constructors

        internal Worksheet(SpireWorksheet worksheet) {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        }

        #endregion Internal Constructors

        #region Private Constructors

        private Worksheet() {
        }

        #endregion Private Constructors

        #region IWorksheet Members

        public string Name {
            get { return _worksheet.Name; }
            set { _worksheet.Name = value; }
        }

        public IEnumerable<IColumn> Columns => _worksheet.Columns.Select(_ => new Column(_));

        public IEnumerable<IRow> Rows => _worksheet.Rows.Select(_ => new Row(_));

        public IColumn CreateColumn(int insertAt = -1) {
            if (insertAt < 0) { insertAt = _worksheet.Columns.Length; }
            _worksheet.InsertColumn(insertAt);
            return new Column(_worksheet.Columns[insertAt]);
        }

        public IRow CreateRow(int insertAt = -1) {
            if (insertAt < 0) { insertAt = _worksheet.Rows.Length; }
            _worksheet.InsertRow(insertAt);
            return new Row(_worksheet.Rows[insertAt]);
        }

        public IRange GetRange(int startColumn, int startRow, int endColumn, int endRow) {
            return new Range(_worksheet[startRow, startColumn, endRow, endColumn]);
        }

        #endregion IWorksheet Members
    }
}