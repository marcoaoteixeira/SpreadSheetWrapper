using System.Collections.Generic;

namespace SpreadSheet.Core {

    public interface IWorksheet {

        #region Properties

        string Name { get; set; }

        IEnumerable<IColumn> Columns { get; }

        IEnumerable<IRow> Rows { get; }

        #endregion

        #region Methods

        IColumn CreateColumn(int insertAt = -1);

        IRow CreateRow(int insertAt = -1);

        IRange GetRange(int startColumn, int startRow, int endColumn, int endRow);

        #endregion Methods
    }
}