namespace SpreadSheet.Core {

    public interface IRow : ICellHolder {

        #region Properties

        int Index { get; }

        int Number { get; }

        #endregion Properties
    }
}