using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class Amex2ProcessorTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void LogueaSiHayCandidatas_O_SiNoEncuentraNinguna()
        {
            string hoja = "AMEX_2";

            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            // Buscar filas candidatas
            var dt = Amex_2Processor.ObtenerFilasCandidatas(archivoPrueba, hoja, null);

            if (dt.Rows.Count > 0)
            {
                // Si encuentra, loguea las filas candidatas
                System.Diagnostics.Debug.WriteLine($"Se encontraron {dt.Rows.Count} filas candidatas para {hoja}:");
                foreach (DataRow row in dt.Rows)
                {
                    // Loguear todos los valores por fila si querés ver todo
                    var valores = string.Join(" | ", row.ItemArray);
                    System.Diagnostics.Debug.WriteLine($"Fila Excel: {row["FilaExcel"]} => {valores}");
                }
            }
            else
            {
                // Si no encuentra, loguea mensaje
                System.Diagnostics.Debug.WriteLine($"No se encontraron filas candidatas para {hoja}.");
            }

            // Procesar SIN eliminar ninguna fila
            List<int> filasAEliminar = new List<int>(); // No elimina ninguna

            int filasContadas;
            double totalBruto = Amex_2Processor.Procesar(archivoPrueba, hoja, filasAEliminar, null, out filasContadas);

            System.Diagnostics.Debug.WriteLine($"Bruto calculado: {totalBruto}, filas contadas: {filasContadas}");

            Assert.IsTrue(totalBruto >= 0, "El total bruto debería ser mayor o igual a cero");
            Assert.IsTrue(filasContadas >= 0, "La cantidad de filas debería ser mayor o igual a cero");
        }
    }
}
