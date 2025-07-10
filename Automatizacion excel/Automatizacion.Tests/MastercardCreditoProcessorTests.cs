using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class MastercardCreditoProcessorTests
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
            string hoja = "Mastercard";

            Assert.IsTrue(File.Exists(archivoTemp), $"No se encontró el archivo: {archivoTemp}");

            var dt = MastercardCreditoProcessor.ObtenerFilasAfectadas(archivoTemp, hoja);

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

            Assert.IsNotNull(dt, "El DataTable no debe ser nulo.");
        }

        [TestMethod]
        public void Procesar_SoloProcesaSinError_OK()
        {
            string hoja = "Mastercard";

            Assert.IsTrue(File.Exists(archivoTemp), $"No se encontró el archivo temporal: {archivoTemp}");

            var dt = MastercardCreditoProcessor.ObtenerFilasAfectadas(archivoTemp, hoja);
            List<int> filasSeleccionadas = new List<int>();
            foreach (DataRow row in dt.Rows)
                filasSeleccionadas.Add(System.Convert.ToInt32(row["FilaExcel"]));

            int cantidadFilas;
            double total = MastercardCreditoProcessor.Procesar(archivoTemp, hoja, filasSeleccionadas, null, out cantidadFilas);

            Assert.IsTrue(total >= 0, "El total sumado debería ser >= 0");
            Assert.IsTrue(cantidadFilas >= 0, "La cantidad de filas válidas debería ser >= 0");

            System.Diagnostics.Debug.WriteLine($"MASTERCARD CRÉDITO - Total columna H: {total}, filas: {cantidadFilas}");
        }
    }
}
