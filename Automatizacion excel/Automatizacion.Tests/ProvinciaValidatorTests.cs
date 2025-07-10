using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using Automatizacion.Core.Formularios.Validaciones;

namespace Automatizacion.Tests
{
    [TestClass]
    public class ProvinciaValidatorTests
    {
        [TestMethod]
        public void Valida_ProvinciaVacia()
        {
            var tabla = new DataTable();
            tabla.Columns.Add("Provincia");
            tabla.Columns.Add("Alicuota");

            tabla.Rows.Add("", "3,00%");

            var errores = ProvinciaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("vacía")));
        }

        [TestMethod]
        public void Valida_AlicuotaInvalida()
        {
            var tabla = new DataTable();
            tabla.Columns.Add("Provincia");
            tabla.Columns.Add("Alicuota");

            tabla.Rows.Add("Buenos Aires", "XXX");

            var errores = ProvinciaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("inválida")));
        }

        [TestMethod]
        public void Valida_Duplicados()
        {
            var tabla = new DataTable();
            tabla.Columns.Add("Provincia");
            tabla.Columns.Add("Alicuota");

            tabla.Rows.Add("Buenos Aires", "3,00%");
            tabla.Rows.Add("Buenos Aires", "3,00%");

            var errores = ProvinciaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("duplicada")));
        }

        [TestMethod]
        public void Valida_RegistroValido()
        {
            var tabla = new DataTable();
            tabla.Columns.Add("Provincia");
            tabla.Columns.Add("Alicuota");

            tabla.Rows.Add("Buenos Aires", "3,00%");

            var errores = ProvinciaValidator.ValidarTabla(tabla);

            Assert.AreEqual(0, errores.Count);
        }
    }
}
