using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using Automatizacion.Core.Formularios.Validaciones;

namespace Automatizacion.Tests
{
    [TestClass]
    public class TasaValidatorTests
    {
        private DataTable GetDataTableBase()
        {
            var tabla = new DataTable();
            tabla.Columns.Add("Cuota");
            tabla.Columns.Add("Codigo_Posnet");
            tabla.Columns.Add("Costo_Visa_Credito");
            tabla.Columns.Add("Costo_American_Express_Credito");
            tabla.Columns.Add("Costo_Mastercard_Credito");
            tabla.Columns.Add("Costo_Argencard_Credito");
            tabla.Columns.Add("Costo_Cabal_Credito");
            tabla.Columns.Add("Costo_Naranja_Credito");
            tabla.Columns.Add("Comision");
            tabla.Columns.Add("IVA");
            tabla.Columns.Add("Comision_mas_IVA");
            return tabla;
        }

        [TestMethod]
        public void Valida_CuotaVacia()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("", "Debito", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "3,19%", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("Cuota vacía")));
        }

        [TestMethod]
        public void Valida_CodigoPosnetVacio()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("0", "", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "3,19%", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("Codigo_Posnet vacío")));
        }

        [TestMethod]
        public void Valida_CostoInvalido()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("0", "Debito", "abc", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "3,19%", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("Costo_Visa_Credito inválido")));
        }

        [TestMethod]
        public void Valida_Costo_Revisar_OK()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("0", "Debito", "Revisar", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "3,19%", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.IsFalse(errores.Exists(e => e.Contains("Costo_Visa_Credito inválido")));
        }

        [TestMethod]
        public void Valida_ComisionInvalida()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("0", "Debito", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "xyz", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.IsTrue(errores.Exists(e => e.Contains("Comision inválido")));
        }

        [TestMethod]
        public void Valida_RegistroValido()
        {
            var tabla = GetDataTableBase();
            tabla.Rows.Add("0", "Debito", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "0,00%", "3,19%", "21,00%", "3,86%");

            var errores = TasaValidator.ValidarTabla(tabla);

            Assert.AreEqual(0, errores.Count);
        }
    }
}
