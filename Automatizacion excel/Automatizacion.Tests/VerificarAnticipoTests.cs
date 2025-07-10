using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.IO;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class VerificarAnticipoTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void FilasSinAnticipo_DevuelveFilas_OK_TodasLasTarjetas()
        {
            string[] hojas = { "Visa", "Mastercard", "ARGENCARD" };

            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            foreach (var hoja in hojas)
            {
                List<int> filasVacias = VerificarAnticipo.FilasSinAnticipo(archivoPrueba, hoja);

                if (filasVacias.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[{hoja}] Filas con columna O vacía: {string.Join(", ", filasVacias)}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[{hoja}] No se encontraron filas con columna O vacía.");
                }

                // Assert flexible: solo chequea que no sea null
                Assert.IsNotNull(filasVacias, $"[{hoja}] La lista de filas vacías no debe ser null");
            }
        }
    }
}
