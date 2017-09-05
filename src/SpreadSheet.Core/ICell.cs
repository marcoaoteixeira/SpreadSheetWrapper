namespace SpreadSheet.Core {

    public interface ICell {

        #region Properties
        
        object Value { get; set; }

        CellValueType ValueType { get; }

        ICellStyle Style { get; }

        CellBorders Borders { get; }

        ICellFormat Format { get; }

        #endregion Properties
    }
}