using System.Drawing;

namespace SpreadSheet.Core {

    public interface ICellBorder {

        #region Properties

        CellBorderPosition Position { get; }

        Color Color { get; set; }

        CellBorderStyle Style { get; set; }

        #endregion Properties
    }
}