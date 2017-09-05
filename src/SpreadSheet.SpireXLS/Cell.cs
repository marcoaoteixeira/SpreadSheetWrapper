using System;
using SpreadSheet.Core;
using SpireCell = Spire.Xls.CellRange;

namespace SpreadSheet.SpireXLS {

    public sealed class Cell : ICell {

        #region Private Fields

        private object _value;
        private CellValueType _valueType;
        private CellBorders _borders;

        #endregion Private Fields

        #region Internal Properties

        internal SpireCell CellImpl { get; }

        #endregion Internal Properties

        #region Public Constructors

        public Cell(SpireCell cell) {
            CellImpl = cell ?? throw new ArgumentNullException(nameof(cell));

            Initialize();
        }

        #endregion Public Constructors

        #region Private Static Methods

        private static CellValueType ParseCellValueType(object value) {
            if (value == null) { return CellValueType.None; }
            if (value.GetType() == typeof(Guid)) { return CellValueType.Text; }
            if (value.GetType() == typeof(TimeSpan)) { return CellValueType.TimeSpan; }

            switch (Type.GetTypeCode(value.GetType())) {
                case TypeCode.Empty:
                case TypeCode.DBNull:
                    return CellValueType.None;

                case TypeCode.Boolean:
                    return CellValueType.TrueFalse;

                case TypeCode.Object:
                case TypeCode.Char:
                case TypeCode.String:
                    return CellValueType.Text;

                case TypeCode.SByte:
                case TypeCode.Byte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return CellValueType.Numeric;

                case TypeCode.DateTime:
                    return CellValueType.DateTime;

                default:
                    throw new InvalidCastException("Unknown value type.");
            }
        }

        #endregion Private Static Methods

        #region Private Methods

        private void Initialize() {
            _borders = new CellBorders {
                Top = new CellBorder(CellImpl.Borders[Spire.Xls.BordersLineType.EdgeTop], CellBorderPosition.Top),
                Right = new CellBorder(CellImpl.Borders[Spire.Xls.BordersLineType.EdgeRight], CellBorderPosition.Right),
                Bottom = new CellBorder(CellImpl.Borders[Spire.Xls.BordersLineType.EdgeBottom], CellBorderPosition.Bottom),
                Left = new CellBorder(CellImpl.Borders[Spire.Xls.BordersLineType.EdgeLeft], CellBorderPosition.Left)
            };
        }

        private void SetValue(object value) {
            _value = value;

            switch (ParseCellValueType(value)) {
                case CellValueType.None:
                    CellImpl.Value = null;
                    break;

                case CellValueType.Text:
                    _value = value.ToString();
                    break;

                case CellValueType.Numeric:
                    CellImpl.NumberValue = Convert.ToDouble(value);
                    break;

                case CellValueType.DateTime:
                    CellImpl.DateTimeValue = Convert.ToDateTime(value);
                    break;

                case CellValueType.TimeSpan:
                    CellImpl.TimeSpanValue = TimeSpan.Parse(value.ToString());
                    break;

                case CellValueType.TrueFalse:
                    CellImpl.BooleanValue = Convert.ToBoolean(value);
                    break;
            }
        }

        private object GetValue() {
            return _value;
        }

        #endregion Private Methods

        #region ICell Members

        public object Value {
            get { return GetValue(); }
            set { SetValue(value); }
        }

        public CellValueType ValueType => ParseCellValueType(Value);

        public ICellStyle Style {
            get { return new CellStyle(CellImpl.Style); }
        }

        public CellBorders Borders {
            get { return _borders; }
        }

        public ICellFormat Format { get; }

        #endregion ICell Members
    }
}