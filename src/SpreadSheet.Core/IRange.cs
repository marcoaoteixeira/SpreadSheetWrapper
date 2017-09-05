namespace SpreadSheet.Core {

    public interface IRange {

        #region Properties

        ICell[][] Cells { get; }

        #endregion Properties

        #region Methods

        void SetStyle(ICellStyle style);

        #endregion
    }
}