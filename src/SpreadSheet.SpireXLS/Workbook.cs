using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SpreadSheet.Core;
using SpireWorkbook = Spire.Xls.Workbook;

namespace SpreadSheet.SpireXLS {

    public sealed class Workbook : IWorkbook, IDisposable {

        #region Private Fields

        private SpireWorkbook _workbook;

        private bool _disposed;

        #endregion Private Fields

        #region Internal Properties

        internal SpireWorkbook WorkbookImpl {
            get { return _workbook; }
        }

        #endregion

        #region Public Constructors

        public Workbook() {
            _workbook = new SpireWorkbook();
            var worksheet = _workbook.Worksheets[0];
        }

        public Workbook(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("Parameter cannot be null, empty or white spaces.", nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException("File not found.", Path.GetFileName(filePath));
            }

            _workbook = new SpireWorkbook();
            _workbook.LoadFromStream(new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read));
        }

        public Workbook(Stream stream) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            _workbook = new SpireWorkbook();
            _workbook.LoadFromStream(stream);
        }

        #endregion Public Constructors

        #region Destructor

        /// <summary>
        /// Destructor
        /// </summary>
        ~Workbook() {
            Dispose(disposing: false);
        }

        #endregion Destructor

        #region Private Methods
        
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> if called from managed code; otherwise <c>false</c>.</param>
        private void Dispose(bool disposing) {
            if (_disposed) { return; }
            if (disposing) {
                // Dispose your managed resources here
                if (_workbook != null) {
                    _workbook.Dispose();
                }
            }

            // Dispose your unmanaged resources here
            _workbook = null;
            _disposed = true;
        }

        private void ThrowIfDisposed() {
            if (_disposed) {
                throw new ObjectDisposedException(nameof(Workbook));
            }
        }

        #endregion Private Methods

        #region IWorkbook Members

        public IEnumerable<IWorksheet> Worksheets {
            get {
                ThrowIfDisposed();

                return _workbook.Worksheets.Select(_ => new Worksheet(_ as Spire.Xls.Worksheet));
            }
        }

        public IWorksheet CreateWorksheet(string name = null, int insertAt = -1) {
            ThrowIfDisposed();

            var worksheet = _workbook.CreateEmptySheet();

            if (!string.IsNullOrWhiteSpace(name)) {
                worksheet.Name = name;
            }

            if (insertAt < 0 || insertAt > _workbook.Worksheets.Count) {
                _workbook.Worksheets.Add(worksheet);
            } else {
                _workbook.Worksheets.Insert(insertAt, worksheet);
            }

            return new Worksheet(worksheet);
        }

        public bool RemoveWorksheet(string name) {
            ThrowIfDisposed();

            var worksheet = _workbook.Worksheets.SingleOrDefault(_ => string.Equals(_.Name, name, StringComparison.CurrentCulture));
            if (worksheet == null) { return false; }

            _workbook.Worksheets.Remove(worksheet);

            return true;
        }

        public void Save() {
            ThrowIfDisposed();

            _workbook.Save();
        }

        public void SaveToFile(string filePath) {
            ThrowIfDisposed();

            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("Parameter cannot be null, empty or white spaces.", nameof(filePath));
            }

            _workbook.SaveToFile(filePath);
        }

        public void SaveToStream(Stream stream) {
            ThrowIfDisposed();

            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            _workbook.SaveToStream(stream);
        }

        #endregion IWorkbook Members

        #region IDisposable Members

        /// <inheritdoc />
        public void Dispose() {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion IDisposable Members
    }
}