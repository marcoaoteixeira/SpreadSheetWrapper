using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadSheet.Core;
using NPOICell = global::NPOI.SS.UserModel.ICell;

namespace SpreadSheet.NPOI {
    public sealed class Cell : ICell {
        #region Internal Properties

        internal  NPOICell CellImpl { get; }

        #endregion

        #region Internal Constructors

        internal Cell(NPOICell cell) {
            CellImpl = cell;
        }

        #endregion

        #region ICell Members
        public object Value { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public CellValueType ValueType => throw new NotImplementedException();

        public ICellStyle Style => throw new NotImplementedException();

        public CellBorders Borders => throw new NotImplementedException();

        public ICellFormat Format => throw new NotImplementedException(); 
        #endregion
    }
}
