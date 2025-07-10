using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class VisaDebitoProcessorTests
    {
        private string archivoPrueba;
        private string hoja = "Visa debito";

        [TestInitialize]
        public void Setup()
        {
            // Hacé una copia del archivo para que el test sea seguro y repetible
            var archivoOriginal = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
            var tempFile = Path.Combine(Path.GetTempPath(), "CONVERSOR_temp_VisaDebito.xlsm");
            File.Copy(archivoOriginal, tempFile, true);
            archivoPrueba = tempFile;
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(archivoPrueba))
                File.Delete(archivoPrueba);
        }

        [TestMethod]
        public void ProcesarVisaDebito_NoExplotaYDevuelveValores()
        {
            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            int filasSumadas;
            double total = VisaDebitoProcessor.Procesar(archivoPrueba, hoja, null, out filasSumadas);

            System.Diagnostics.Debug.WriteLine($"Total bruto calculado: {total}");
            System.Diagnostics.Debug.WriteLine($"Filas sumadas: {filasSumadas}");

            Assert.IsTrue(total >= 0, "El total bruto debe ser >= 0");
            Assert.IsTrue(filasSumadas >= 0, "La cantidad de filas debe ser >= 0");
        }
    }
}
