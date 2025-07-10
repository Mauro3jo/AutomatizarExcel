using Microsoft.VisualStudio.TestTools.UnitTesting;
using Automatizacion_excel.Paso1;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Automatizacion.Tests.Excel
{
    [TestClass]
    public class ControlTasasTests
    {
        private string archivoPrueba;

        [TestInitialize]
        public void Setup()
        {
            archivoPrueba = Path.GetFullPath(Path.Combine("TestFiles", "CONVERSOR.xlsm"));
        }

        [TestMethod]
        public void ObtenerTasasDesdeBD_DevuelveTasasParaLasTresTarjetas()
        {
            var tarjetas = new[] { "Visa", "Mastercard", "ARGENCARD" };

            foreach (var tarjeta in tarjetas)
            {
                var tasas = ControlTasas.ObtenerTasasDesdeBD(tarjeta);
                Assert.IsNotNull(tasas, $"El diccionario de tasas para {tarjeta} no debe ser nulo");
                Assert.IsTrue(tasas.Count > 0, $"Debe haber al menos una cuota con tasa para {tarjeta}");

                System.Diagnostics.Debug.WriteLine($"{tarjeta}: " + string.Join("; ", tasas.Select(kvp => $"{kvp.Key}:{kvp.Value}")));
            }
        }

        [TestMethod]
        public void VerificarExcesos_ParaCadaTarjeta_Excel_OK()
        {
            var tarjetas = new[]
            {
                new { Tarjeta = "Visa", Hoja = "Visa" },
                new { Tarjeta = "Mastercard", Hoja = "Mastercard" },
                new { Tarjeta = "ARGENCARD", Hoja = "ARGENCARD" }
            };

            Assert.IsTrue(File.Exists(archivoPrueba), $"No se encontró el archivo: {archivoPrueba}");

            foreach (var t in tarjetas)
            {
                var tasas = ControlTasas.ObtenerTasasDesdeBD(t.Tarjeta);
                Assert.IsNotNull(tasas, $"El diccionario de tasas para {t.Tarjeta} no debe ser nulo");

                var filasConExceso = ControlTasas.VerificarExcesos(archivoPrueba, t.Hoja, tasas);

                // Test mínimo: que devuelva algo (puede ser vacío)
                Assert.IsNotNull(filasConExceso, $"La lista de filas con exceso para {t.Tarjeta} no debe ser nula");

                System.Diagnostics.Debug.WriteLine($"{t.Tarjeta} ({t.Hoja}): Filas con exceso = {string.Join(", ", filasConExceso)}");
            }
        }
    }
}
