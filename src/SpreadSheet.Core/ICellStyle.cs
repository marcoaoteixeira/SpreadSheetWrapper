using System;
using System.Drawing;

namespace SpreadSheet.Core {

    public interface ICellStyle {

        #region Properties

        string FontFamily { get; set; }

        double FontSize { get; set; }

        bool Bold { get; set; }

        bool Underline { get; set; }

        bool Italic { get; set; }

        Color ForeColor { get; set; }

        Color BackColor { get; set; }

        #endregion Properties

        #region Methods

        void Clone(ICellStyle style);

        #endregion Methods
    }

    /// <summary>
    /// Singleton Pattern implementation for NormalCellStyle. (see: https://en.wikipedia.org/wiki/Singleton_pattern)
    /// </summary>
    public sealed class NormalCellStyle : ICellStyle {

        #region Private Static Read-Only Fields

        private static readonly NormalCellStyle _instance = new NormalCellStyle();

        #endregion Private Static Read-Only Fields

        #region Public Static Properties

        /// <summary>
        /// Gets the unique instance of NormalCellStyle.
        /// </summary>
        public static NormalCellStyle Instance {
            get { return _instance; }
        }

        #endregion Public Static Properties

        #region Static Constructors

        // Explicit static constructor to tell the C# compiler
        // not to mark type as beforefieldinit
        static NormalCellStyle() {
        }

        #endregion Static Constructors

        #region Private Constructors

        private NormalCellStyle() {
        }

        #endregion Private Constructors

        #region ICellStyle Members

        public string FontFamily {
            get { return "Arial"; }
            set { /* DO NOTHING */ }
        }

        public double FontSize {
            get { return 12d; }
            set { /* DO NOTHING */ }
        }

        public bool Bold {
            get { return false; }
            set { /* DO NOTHING */ }
        }

        public bool Underline {
            get { return false; }
            set { /* DO NOTHING */ }
        }

        public bool Italic {
            get { return false; }
            set { /* DO NOTHING */ }
        }

        public Color ForeColor {
            get { return Color.Black; }
            set { /* DO NOTHING */ }
        }

        public Color BackColor {
            get { return Color.Transparent; }
            set { /* DO NOTHING */ }
        }

        public void Clone(ICellStyle style) {
        }

        #endregion ICellStyle Members
    }
}