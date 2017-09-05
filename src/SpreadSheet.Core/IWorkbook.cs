using System;
using System.Collections.Generic;
using System.IO;

namespace SpreadSheet.Core {

    public interface IWorkbook {

        #region Properties

        IEnumerable<IWorksheet> Worksheets { get; }

        #endregion Properties

        #region Methods

        IWorksheet CreateWorksheet(string name = null, int insertAt = -1);

        bool RemoveWorksheet(string name);

        void Save();
        
        void SaveToFile(string filePath);

        void SaveToStream(Stream stream);

        #endregion Methods
    }
}