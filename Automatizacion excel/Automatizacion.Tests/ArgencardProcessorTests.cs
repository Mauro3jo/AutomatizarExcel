using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Automatizacion.Tests
{
    [TestClass]
    public class ArgencardProcessorTests
    {
        private string archivoPrueba;
        private string archivoTemp;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
            archivoTemp = Path.Combine(Path.GetTempPath(), $"CONVERSOR_{System.Guid.NewGuid()}.xlsm");
            File.Copy(archivoPrueba, archivoTemp, overwrite: true);
        }

        [TestCleanup]
        public void Cleanup()
        {
            if (File.Exists(archivoTemp))
                File.Delete(archivoTemp);
        }

        [TestMethod]
        public void ObtenerFilasAfectadas_LogueaFilas_OK()
        {
            string hoja = "ARGENCARD";

            Assert.IsTrue(File.Exists(archivoTemp), $"No se encontró el archivo temporal: {archivoTemp}");

            // Vista previa: buscar filas afectadas
            var dt = ArgencardProcessor.ObtenerFilasAfectadas(archivoTemp, hoja, null);

            if (dt.Rows.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"Se encontraron {dt.Rows.Count} filas afectadas para {hoja}:");
                foreach (DataRow row in dt.Rows)
                {
                    var valores = string.Join(" | ", row.ItemArray);
                    System.Diagnostics.Debug.WriteLine($"Fila Excel: {row["FilaExcel"]} => {valores}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"No se encontraron filas afectadas para {hoja}.");
            }

            // Test mínimo: no explota y devuelve DataTable (puede estar vacío)
            Assert.IsNotNull(dt, "El DataTable no debe ser nulo.");
        }

        [TestMethod]
        public void Procesar_SoloProcesaSinError_OK()
        {
            string hoja = "ARGENCARD";

            Assert.IsTrue(File.Exists(archivoTemp), $"No se encontró el archivo temporal: {archivoTemp}");

            // Simula selección de todas las filas afectadas
            var dt = ArgencardProcessor.ObtenerFilasAfectadas(archivoTemp, hoja, null);
            List<int> filasSeleccionadas = new List<int>();
            foreach (DataRow row in dt.Rows)
                filasSeleccionadas.Add(System.Convert.ToInt32(row["FilaExcel"]));

            int filasContadas;
            double total = ArgencardProcessor.Procesar(archivoTemp, hoja, filasSeleccionadas, null, out filasContadas);

            Assert.IsTrue(total >= 0, "El total sumado debería ser >= 0");
            Assert.IsTrue(filasContadas >= 0, "La cantidad de filas válidas debería ser >= 0");

            System.Diagnostics.Debug.WriteLine($"ARGENCARD - Total columna H: {total}, filas: {filasContadas}");
        }
    }
}
