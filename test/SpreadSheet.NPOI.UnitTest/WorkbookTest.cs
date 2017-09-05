using System.IO;
using System.Linq;
using SpreadSheet.NPOI.UnitTest.Resources;
using Xunit;

namespace SpreadSheet.NPOI.UnitTest {

    public class WorkbookTest : TestBase {
        private static readonly string TempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetTempFileName()));

        protected override void SetUp() {
            File.WriteAllBytes(string.Concat(TempFilePath, Utils.XLSX_EXT), Resource.EMPTY_WORKBOOK_XLSX);
            File.WriteAllBytes(string.Concat(TempFilePath, Utils.XLS_EXT), Resource.EMPTY_WORKBOOK_XLS);
        }

        protected override void TearDown() {
            File.Delete(string.Concat(TempFilePath, Utils.XLSX_EXT));
            File.Delete(string.Concat(TempFilePath, Utils.XLS_EXT));
        }

        [Theory]
        [InlineData(Utils.XLSX_EXT)]
        [InlineData(Utils.XLS_EXT)]
        public void Open_Workbook(string extension) {
            var path = string.Concat(TempFilePath, extension);

            using (new Workbook(path)) { }
        }

        [Theory]
        [InlineData(Utils.XLSX_EXT)]
        [InlineData(Utils.XLS_EXT)]
        public void Get_Sheets_From_EmptyWorkbook(string extension) {
            var path = string.Concat(TempFilePath, extension);

            using (var workbook = new Workbook(path)) {
                Assert.Equal(1, workbook.Worksheets.Count());
            }
        }

        [Theory]
        [InlineData(Utils.XLSX_EXT)]
        [InlineData(Utils.XLS_EXT)]
        public void Create_Sheet_In_EmptyWorkbook(string extension) {
            var path = string.Concat(TempFilePath, extension);

            using (var workbook = new Workbook(path)) {
                workbook.CreateWorksheet();

                Assert.Equal(2, workbook.Worksheets.Count());
            }
        }

        [Theory]
        [InlineData(Utils.XLSX_EXT)]
        [InlineData(Utils.XLS_EXT)]
        public void Remove_Sheet_From_EmptyWorkbook(string extension) {
            var path = string.Concat(TempFilePath, extension);

            using (var workbook = new Workbook(path)) {
                workbook.RemoveWorksheet("Plan1");

                Assert.Equal(0, workbook.Worksheets.Count());
            }
        }

        [Theory]
        [InlineData(Utils.XLSX_EXT)]
        [InlineData(Utils.XLS_EXT)]
        public void SaveToFile_EmptyWorkbook(string extension) {
            var path = string.Concat(TempFilePath, extension);
            var newPath = string.Concat(TempFilePath, "_NEW_", extension);

            using (var workbook = new Workbook(path)) {
                workbook.SaveToFile(newPath);

                Assert.True(File.Exists(newPath));

                File.Delete(newPath);
            }
        }
    }
}