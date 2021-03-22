using ExcelEnt.Tests.Binders;
using ExcelEnt.Tests.ExpectedData;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace ExcelEnt.Tests
{
    [TestClass]
    public class ExcelEntBindTest
    {
        /// <summary>
        /// Test binding all rows from file
        /// </summary>
        [TestMethod]
        public void TestFullBind()
        {
            var binder = new FullBinder();
            var entities = binder.Bind(Expected.TestItemsPath);
            var expected = Expected.Data;

            CollectionAssert.AreEqual(entities, expected);
        }

        /// <summary>
        /// Test binding some rows from file
        /// </summary>
        [TestMethod]
        public void TestPartBind()
        {
            var skipRows = 1;
            var takeRows = 2;

            var binder = new PartBinder();
            var entities = binder.Bind(Expected.TestItemsPath, skipRows, takeRows);
            var expected = Expected.Data.Skip(skipRows).Take(takeRows).ToArray();

            CollectionAssert.AreEqual(entities, expected);
        }
    }
}
