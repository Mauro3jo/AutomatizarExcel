using System.Collections.Generic;
using System.Data;

namespace Automatizacion.Core.Formularios.Validaciones
{
    public static class ProvinciaValidator
    {
        public static List<string> ValidarTabla(DataTable tabla)
        {
            var errores = new List<string>();
            var nombresProvincias = new HashSet<string>(System.StringComparer.InvariantCultureIgnoreCase);

            foreach (DataRow row in tabla.Rows)
            {
                var provincia = row["Provincia"]?.ToString()?.Trim();
                var alicuotaStr = row["Alicuota"]?.ToString()?.Trim();

                if (string.IsNullOrWhiteSpace(provincia))
                {
                    errores.Add("Provincia vacía o nula en una fila.");
                }
                else if (!nombresProvincias.Add(provincia))
                {
                    errores.Add($"Provincia duplicada: {provincia}");
                }

                if (string.IsNullOrWhiteSpace(alicuotaStr))
                {
                    errores.Add($"Alicuota vacía para la provincia: {provincia}");
                }
                else
                {
                    // Validar formato "n,nn%" o "n,nn %"
                    var numeroStr = alicuotaStr.Replace("%", "").Replace(",", ".").Trim();
                    if (!decimal.TryParse(numeroStr, out var alicuota) || alicuota < 0)
                        errores.Add($"Alicuota inválida ('{row["Alicuota"]}') para provincia: {provincia}");
                }
            }
            return errores;
        }
    }
}
