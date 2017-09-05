using System;
using System.Collections.Generic;
using System.IO;
using SpreadSheet.Core;
using NPOIWorkbook = global::NPOI.SS.UserModel.IWorkbook;

namespace SpreadSheet.NPOI {

    public sealed class Workbook : IWorkbook, IDisposable {

        #region Private Constants

        private const string WORKSHEET_NAME_PATTERN = "Sheet{0}";

        #endregion Private Constants

        #region Private Read-Only Fields

        private readonly string _filePath;
        private readonly Stream _stream;

        #endregion Private Read-Only Fields

        #region Private Fields

        private NPOIWorkbook _workbook;
        private bool _disposed;

        #endregion Private Fields

        #region Internal Properties

        internal NPOIWorkbook WorkbookImpl {
            get { return _workbook; }
        }

        #endregion Internal Properties

        #region Public Constructors

        public Workbook(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("Parameter cannot be null, empty or white spaces.", nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException("File not found.", fileName: Path.GetFileName(filePath));
            }

            _filePath = filePath;
            _workbook = global::NPOI.SS.UserModel.WorkbookFactory.Create(filePath);
        }

        public Workbook(Stream stream) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _workbook = global::NPOI.SS.UserModel.WorkbookFactory.Create(stream);
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

        private void Dispose(bool disposing) {
            if (_disposed) { return; }
            if (disposing) {
                // Dispose your managed resources here
                if (_workbook != null) {
                    _workbook.Close();
                    //_workbook.Dispose(); Not implemented on NPOI (???)
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

        private IEnumerable<IWorksheet> GetWorksheets() {
            for (var idx = 0; idx < WorkbookImpl.NumberOfSheets; idx++) {
                yield return new Worksheet(WorkbookImpl.GetSheetAt(idx));
            }
        }

        #endregion Private Methods

        #region IWorkbook Members

        public IEnumerable<IWorksheet> Worksheets => GetWorksheets();

        public IWorksheet CreateWorksheet(string name = null, int insertAt = -1) {
            var sheetName = string.IsNullOrWhiteSpace(name)
                ? string.Format(WORKSHEET_NAME_PATTERN, WorkbookImpl.NumberOfSheets + 1)
                : name;

            var sheet = WorkbookImpl.CreateSheet(sheetName);

            if (insertAt >= 0) {
                WorkbookImpl.SetSheetOrder(sheetName, insertAt);
            }

            return new Worksheet(sheet);
        }

        public bool RemoveWorksheet(string name) {
            if (string.IsNullOrWhiteSpace(name)) { return false; }

            var sheetIndex = WorkbookImpl.GetSheetIndex(name);

            if (sheetIndex < 0) { return false; }

            WorkbookImpl.RemoveSheetAt(sheetIndex);

            return true;
        }

        public void Save() {
            if (!string.IsNullOrWhiteSpace(_filePath)) { SaveToFile(_filePath); } else { SaveToStream(_stream); }
        }

        public void SaveToFile(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                throw new ArgumentException("Parameter cannot be null, empty or white spaces.", nameof(filePath));
            }

            using (var stream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.Write)) {
                WorkbookImpl.Write(stream);
            }
        }

        public void SaveToStream(Stream stream) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            WorkbookImpl.Write(stream);
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