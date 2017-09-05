using System;
using System.Linq;
using SpreadSheet.Core;
using SpireRange = Spire.Xls.CellRange;

namespace SpreadSheet.SpireXLS {

    public sealed class Range : IRange {

        #region Private Read-Only Fields

        private readonly SpireRange _range;

        #endregion Private Read-Only Fields

        #region Private Fields

        private ICell[][] _cache;

        #endregion Private Fields

        #region Internal Properties

        internal SpireRange RangeImpl {
            get { return _range; }
        }

        #endregion Internal Properties

        #region Internal Constructors

        internal Range(SpireRange range) {
            _range = range ?? throw new ArgumentNullException(nameof(range));

            Initialize();
        }

        #endregion Internal Constructors

        #region Private Constructors

        private Range() {
        }

        #endregion Private Constructors

        #region Private Methods

        private void Initialize() {
            _cache = new Cell[_range.Cells.Length][];
            for (var idx = 0; idx < _cache.Length; idx++) {
                _cache[idx] = _range.Cells[idx].Cells.Select(_ => new Cell(_)).ToArray();
            }
        }

        #endregion Private Methods

        #region IRange Members

        public ICell[][] Cells => _cache;

        public void SetStyle(ICellStyle style) {
            foreach (var cells in _cache) {
                foreach (var cell in cells) {
                    cell.Style.Clone(style);
                }
            }
        }

        #endregion IRange Members
    }
}