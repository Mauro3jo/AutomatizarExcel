using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion.Core.Excel.Servicios;
using System.Collections.Generic;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class ExcelReaderServiceTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.Combine("TestFiles", "Ejemplo.xlsx");
        }

        [TestMethod]
        public void ObtenerNombresHojas_DeberiaRetornarHojasCorrectas()
        {
            var reader = new ExcelReaderService();

            List<string> hojas = reader.ObtenerNombresHojas(archivoPrueba);

            Assert.IsTrue(hojas.Contains("Ventas"));
            Assert.IsTrue(hojas.Contains("Compras"));
        }
    }
}
