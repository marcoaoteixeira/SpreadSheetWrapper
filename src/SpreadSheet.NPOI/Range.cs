using System;
using SpreadSheet.Core;

namespace SpreadSheet.NPOI {

    public class Range : IRange {

        #region Public Constructors

        public Range(ICell[][] range) {
            Cells = range ?? throw new ArgumentNullException(nameof(range));
        }

        #endregion Public Constructors

        #region IRange Members

        public ICell[][] Cells { get; }

        public void SetStyle(ICellStyle style) {
            throw new NotImplementedException();
        }

        #endregion IRange Members
    }
}