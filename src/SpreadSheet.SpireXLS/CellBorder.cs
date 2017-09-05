using System;
using System.Drawing;
using Spire.Xls;
using SpreadSheet.Core;
using SpireCellBorder = Spire.Xls.Core.IBorder;

namespace SpreadSheet.SpireXLS {

    public sealed class CellBorder : ICellBorder {

        #region Private Read-Only Fields

        private readonly CellBorderPosition _position;

        #endregion Private Read-Only Fields

        #region Internal Properties

        internal SpireCellBorder BorderImpl { get; }

        #endregion Internal Properties

        #region Internal Constructors

        internal CellBorder(SpireCellBorder border, CellBorderPosition position) {
            BorderImpl = border ?? throw new ArgumentNullException(nameof(border));
            _position = position;
        }

        #endregion Internal Constructors

        #region Private Constructors

        private CellBorder() {
        }

        #endregion Private Constructors

        #region ICellBorder Members

        public CellBorderPosition Position {
            get { return _position; }
        }

        public Color Color {
            get { return BorderImpl.Color; }
            set { BorderImpl.Color = value; }
        }

        public CellBorderStyle Style {
            get { return (CellBorderStyle)BorderImpl.LineStyle; }
            set { BorderImpl.LineStyle = (LineStyleType)value; }
        }

        #endregion ICellBorder Members
    }
}