using Aspose.Cells;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace AsposeExtensions.Tests
{
    [TestClass]
    public class AsposeExtensionsTest
    {
        [TestMethod]
        [DeploymentItem("Files\\Contas-a-Pagar.xlsx")]
        [ExpectedException(typeof(AsposeExtensionsSheetNotFoundException))]
        public void RowToModel_SheetNotFound()
        {
            var path = string.Concat("Contas-a-Pagar.xlsx");
            var excel = new Workbook(path);
            var row = excel.RowToModel<ContasAPagarSheetRow>("whatever name", 1);
        }

        [TestMethod]
        [DeploymentItem("Files\\Contas-a-Pagar.xlsx")]
        public void RowToModel_SheetRowLoaded()
        {
            var path = string.Concat("Contas-a-Pagar.xlsx");
            var excel = new Workbook(path);
            var row = excel.RowToModel<ContasAPagarSheetRow>("Contas a Pagar", 5, 2);

            Assert.AreEqual(row.SheetName, "Contas a Pagar");
            Assert.AreEqual(row.Row, 5);
        }

        [TestMethod]
        [DeploymentItem("Files\\Contas-a-Pagar.xlsx")]
        public void RowsToModelList_SheetLoaded40Rows()
        {
            var path = string.Concat("Contas-a-Pagar.xlsx");
            var excel = new Workbook(path);
            var rows = excel.RowsToModelList<ContasAPagarSheetRow>("Contas a Pagar", 2);

            Assert.AreEqual(rows.Count(), 40);
        }
    }
}
