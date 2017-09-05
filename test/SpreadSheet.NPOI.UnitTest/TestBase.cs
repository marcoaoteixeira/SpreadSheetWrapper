using System;

namespace SpreadSheet.NPOI.UnitTest {

    public abstract class TestBase : IDisposable {

        protected TestBase() {
            SetUp();
        }

        protected virtual void SetUp() {
        }

        protected virtual void TearDown() {
        }

        public void Dispose() {
            TearDown();
        }
    }
}