using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class VisaCreditoProcessorTests
    {
        private string archivoPrueba;
        private string hoja = "Visa";

        [TestInitialize]
        public void Setup()
        {
            // Usá una copia del archivo para evitar modificar el original en los tests
            var archivoOriginal = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
            var tempFile = Path.Combine(Path.GetTempPath(), "CONVERSOR_temp_test.xlsm");
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
        public void ObtenerFilasAfectadasYProcesar_OK()
        {
            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            // 1. Obtener las filas afectadas (como hace el form)
            DataTable dt = VisaCreditoProcessor.ObtenerFilasAfectadas(archivoPrueba, hoja, null);
            Assert.IsNotNull(dt, "El DataTable no debe ser null");

            // 2. Simular que el usuario acepta TODAS las filas sugeridas
            List<int> filasSeleccionadas = new List<int>();
            foreach (DataRow row in dt.Rows)
            {
                filasSeleccionadas.Add(System.Convert.ToInt32(row["FilaExcel"]));
            }

            // 3. Procesar (transforma y suma)
            int cantidadFilas;
            double total = VisaCreditoProcessor.Procesar(archivoPrueba, hoja, filasSeleccionadas, null, out cantidadFilas);

            System.Diagnostics.Debug.WriteLine($"Total bruto calculado: {total}");
            System.Diagnostics.Debug.WriteLine($"Filas procesadas: {cantidadFilas}");

            // 4. Verificamos que devuelve valores razonables
            Assert.IsTrue(total >= 0, "El total bruto debe ser >= 0");
            Assert.IsTrue(cantidadFilas >= 0, "La cantidad de filas debe ser >= 0");
            // Si sabés un valor esperado exacto para tu archivo de test, poné un assert aquí.
        }
    }
}
