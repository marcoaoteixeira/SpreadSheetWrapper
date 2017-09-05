using System.Collections.Generic;

namespace SpreadSheet.Core {

    public interface ICellHolder {

        #region Properties

        IEnumerable<ICell> Cells { get; }

        #endregion Properties

        #region Methods

        ICell CreateCell(int insertAt = -1);

        #endregion Methods
    }
}