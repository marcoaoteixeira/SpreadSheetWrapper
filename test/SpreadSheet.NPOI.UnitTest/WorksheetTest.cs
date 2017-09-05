using System.IO;
using System.Linq;
using SpreadSheet.NPOI.UnitTest.Resources;
using Xunit;

namespace SpreadSheet.NPOI.UnitTest {

    public class WorksheetTest : TestBase {
        private static readonly string TempFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetTempFileName()));

        protected override void SetUp() {
            File.WriteAllBytes(string.Concat(TempFilePath, Utils.XLSX_EXT), Resource.WORKBOOK_XLSX);
            File.WriteAllBytes(string.Concat(TempFilePath, Utils.XLS_EXT), Resource.WORKBOOK_XLS);
        }

        protected override void TearDown() {
            File.Delete(string.Concat(TempFilePath, Utils.XLSX_EXT));
            File.Delete(string.Concat(TempFilePath, Utils.XLS_EXT));
        }

        [Theory]
        [InlineData(Utils.XLS_EXT)]
        [InlineData(Utils.XLSX_EXT)]
        public void Change_Worksheet_Name(string extension) {
            var path = string.Concat(TempFilePath, extension);
            var oldName = string.Empty;
            var newName = "CHANGED";

            using (var workbook = new Workbook(path)) {
                var sheet = workbook.Worksheets.First();

                oldName = sheet.Name;
                sheet.Name = newName;

                Assert.NotEqual(oldName, sheet.Name);
            }
        }

        [Theory]
        [InlineData(Utils.XLS_EXT)]
        [InlineData(Utils.XLSX_EXT)]
        public void Get_Range_From_Worksheet(string extension) {
            var path = string.Concat(TempFilePath, extension);

            using (var workbook = new Workbook(path)) {
                var sheet = workbook.Worksheets.First();

                var range = sheet.GetRange(startColumn: 0, startRow: 0, endColumn: 9, endRow: 9);

                Assert.NotNull(range);
            }
        }
    }
}