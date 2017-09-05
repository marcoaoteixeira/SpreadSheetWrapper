using System;
using System.Collections.Generic;
using SpreadSheet.Core;
using SpireRow = Spire.Xls.CellRange;

namespace SpreadSheet.SpireXLS {

    public sealed class Row : IRow {

        #region Private Read-Only Fields

        private readonly SpireRow _row;

        #endregion Private Read-Only Fields

        #region Internal Properties

        internal SpireRow RowImpl {
            get { return _row; }
        }

        #endregion Internal Properties

        #region Internal Constructors

        internal Row(SpireRow row) {
            _row = row ?? throw new ArgumentNullException(nameof(row));
        }

        #endregion Internal Constructors

        #region Private Constructors

        private Row() {
        }

        #endregion Private Constructors

        #region IColumn Members

        public int Index => _row.Row;

        public int Number => _row.Row + 1;

        public IEnumerable<ICell> Cells => throw new NotImplementedException();

        public ICell CreateCell(int insertAt = -1) {
            throw new NotImplementedException();
        }

        #endregion IColumn Members
    }
}