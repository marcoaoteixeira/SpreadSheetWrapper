using System;
using System.Drawing;
using SpreadSheet.Core;
using SpireCellStyle = Spire.Xls.CellStyle;

namespace SpreadSheet.SpireXLS {

    public sealed class CellStyle : ICellStyle {

        #region Private Read-Only Fields

        private readonly SpireCellStyle _cellStyle;

        #endregion Private Read-Only Fields

        #region Internal Properties

        internal SpireCellStyle CellStyleImpl {
            get { return _cellStyle; }
        }

        #endregion Internal Properties

        #region Internal Constructors

        internal CellStyle(SpireCellStyle cellStyle) {
            _cellStyle = cellStyle ?? throw new ArgumentNullException(nameof(cellStyle));
        }

        #endregion Internal Constructors

        #region Private Constructors

        private CellStyle() {
        }

        #endregion Private Constructors

        #region ICellStyle Members

        public string FontFamily {
            get { return _cellStyle.Font.FontName; }
            set { _cellStyle.Font.FontName = value; }
        }

        public double FontSize {
            get { return _cellStyle.Font.Size; }
            set { _cellStyle.Font.Size = value; }
        }

        public bool Bold {
            get { return _cellStyle.Font.IsBold; }
            set { _cellStyle.Font.IsBold = value; }
        }

        public bool Underline {
            get { return _cellStyle.Font.Underline != Spire.Xls.FontUnderlineType.None; }
            set { _cellStyle.Font.Underline = value ? Spire.Xls.FontUnderlineType.Single : Spire.Xls.FontUnderlineType.None; }
        }

        public bool Italic {
            get { return _cellStyle.Font.IsItalic; }
            set { _cellStyle.Font.IsItalic = value; }
        }

        public Color ForeColor {
            get { return _cellStyle.Font.Color; }
            set { _cellStyle.Font.Color = value; }
        }

        public Color BackColor {
            get { return _cellStyle.Color; }
            set { _cellStyle.Color = value; }
        }

        public void Clone(ICellStyle style) {
            var innerStyle = style ?? NormalCellStyle.Instance;

            FontFamily = innerStyle.FontFamily;
            FontSize = innerStyle.FontSize;
            Bold = innerStyle.Bold;
            Underline = innerStyle.Underline;
            Italic = innerStyle.Italic;
            ForeColor = innerStyle.ForeColor;
            BackColor = innerStyle.BackColor;
        }

        #endregion ICellStyle Members
    }
}