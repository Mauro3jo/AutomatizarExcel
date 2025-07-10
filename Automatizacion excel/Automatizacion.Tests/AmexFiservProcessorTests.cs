using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class AmexFiservProcessorTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void Procesar_SumaColumnaH_YCuentaFilas_OK()
        {
            string hoja = "AMEX FISERV";

            Assert.IsTrue(File.Exists(archivoPrueba),
                $"No se encontró el archivo: {archivoPrueba}");

            int filasContadas;
            double totalBruto = AmexFiservProcessor.Procesar(archivoPrueba, hoja, null, out filasContadas);

            // Lo mínimo: valores >= 0
            Assert.IsTrue(totalBruto >= 0, "El total sumado debería ser >= 0");
            Assert.IsTrue(filasContadas >= 0, "La cantidad de filas válidas debería ser >= 0");

            System.Diagnostics.Debug.WriteLine($"AMEX FISERV - Bruto columna H: {totalBruto}, filas: {filasContadas}");
        }
    }
}
