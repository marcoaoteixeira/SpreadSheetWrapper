namespace SpreadSheet.Core {

    public sealed class CellBorders {

        #region Private Read-Only Fields

        private readonly ICellBorder[] _cache = new ICellBorder[4];

        #endregion Private Read-Only Fields

        #region Public Properties

        public ICellBorder Top {
            get { return _cache[0]; }
            set { _cache[0] = value; }
        }

        public ICellBorder Right {
            get { return _cache[1]; }
            set { _cache[1] = value; }
        }

        public ICellBorder Bottom {
            get { return _cache[2]; }
            set { _cache[2] = value; }
        }

        public ICellBorder Left {
            get { return _cache[3]; }
            set { _cache[3] = value; }
        }

        #endregion Public Properties
    }
}