namespace SpreadSheet.Core {

    public interface IColumn : ICellHolder {

        #region Properties

        int Index { get; }

        string Name { get; }

        #endregion Properties
    }
}