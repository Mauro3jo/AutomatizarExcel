using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class MastercardDebitoProcessorTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void Procesar_SumaColumnaH_OK()
        {
            string hoja = "Mastercard debito";

            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            int filasSumadas;
            double total = MastercardDebitoProcessor.Procesar(archivoPrueba, hoja, null, out filasSumadas);

            Assert.IsTrue(total >= 0, "El total sumado debería ser >= 0");
            Assert.IsTrue(filasSumadas >= 0, "La cantidad de filas válidas debería ser >= 0");

            System.Diagnostics.Debug.WriteLine($"MASTERCARD DEBITO - Total columna H: {total}, filas: {filasSumadas}");
        }
    }
}
