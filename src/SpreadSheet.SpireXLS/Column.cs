using System;
using System.Collections.Generic;
using SpreadSheet.Core;
using SpireColumn = Spire.Xls.CellRange;

namespace SpreadSheet.SpireXLS {

    public sealed class Column : IColumn {

        #region Private Read-Only Fields

        private readonly SpireColumn _column;

        #endregion Private Read-Only Fields

        #region Internal Properties

        internal SpireColumn ColumnImpl {
            get { return _column; }
        }

        #endregion Internal Properties

        #region Internal Constructors

        internal Column(SpireColumn column) {
            _column = column ?? throw new ArgumentNullException(nameof(column));
        }

        #endregion Public Constructors

        #region Private Constructors

        private Column() {
        }

        #endregion Private Constructors

        #region IColumn Members

        public int Index => _column.Column;

        public string Name => _column.CellStyleName;

        public IEnumerable<ICell> Cells => throw new NotImplementedException();

        public ICell CreateCell(int insertAt = -1) {
            throw new NotImplementedException();
        }

        #endregion IColumn Members
    }
}