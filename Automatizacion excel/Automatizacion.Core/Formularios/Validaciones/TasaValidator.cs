using System.Collections.Generic;
using System.Data;

namespace Automatizacion.Core.Formularios.Validaciones
{
    public static class TasaValidator
    {
        public static List<string> ValidarTabla(DataTable tabla)
        {
            var errores = new List<string>();
            var nombresCuotas = new HashSet<string>(System.StringComparer.InvariantCultureIgnoreCase);

            foreach (DataRow row in tabla.Rows)
            {
                var cuota = row["Cuota"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(cuota))
                    errores.Add("Cuota vacía o nula.");

                var codigoPosnet = row["Codigo_Posnet"]?.ToString()?.Trim();
                if (string.IsNullOrWhiteSpace(codigoPosnet))
                    errores.Add($"Codigo_Posnet vacío para Cuota: {cuota}");

                // Validar todos los campos de costo de tarjeta
                string[] costos = {
                    "Costo_Visa_Credito", "Costo_American_Express_Credito",
                    "Costo_Mastercard_Credito", "Costo_Argencard_Credito",
                    "Costo_Cabal_Credito", "Costo_Naranja_Credito"
                };
                foreach (var campo in costos)
                {
                    var valor = row[campo]?.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(valor) && valor != "Revisar")
                    {
                        var numeroStr = valor.Replace("%", "").Replace(",", ".").Trim();
                        if (!decimal.TryParse(numeroStr, out var numero) || numero < 0)
                            errores.Add($"{campo} inválido ('{valor}') para Cuota: {cuota}");
                    }
                }

                // Validar Comision, IVA, Comision_mas_IVA
                string[] extras = { "Comision", "IVA", "Comision_mas_IVA" };
                foreach (var campo in extras)
                {
                    var valor = row[campo]?.ToString()?.Trim();
                    var numeroStr = valor?.Replace("%", "").Replace(",", ".").Trim();
                    if (!string.IsNullOrWhiteSpace(numeroStr) &&
                        (!decimal.TryParse(numeroStr, out var numero) || numero < 0))
                        errores.Add($"{campo} inválido ('{valor}') para Cuota: {cuota}");
                }
            }
            return errores;
        }
    }
}
