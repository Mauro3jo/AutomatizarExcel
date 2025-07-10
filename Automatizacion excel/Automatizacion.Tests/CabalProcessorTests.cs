using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class CabalProcessorTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void Procesar_SumaColumnaF_YCuentaFilas_OK()
        {
            string hoja = "CABAL";

            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            int filasContadas;
            double totalBruto = CabalProcessor.Procesar(archivoPrueba, hoja, null, out filasContadas);

            Assert.IsTrue(totalBruto >= 0, "El total bruto debería ser mayor o igual a cero");
            Assert.IsTrue(filasContadas >= 0, "La cantidad de filas válidas debería ser mayor o igual a cero");

            System.Diagnostics.Debug.WriteLine($"CABAL - Bruto columna F: {totalBruto}, filas: {filasContadas}");
        }
    }
}
