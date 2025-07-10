using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace Automatizacion_excel.Paso4
{
    internal static class IIBBHelper
    {
        /// <summary>
        /// Devuelve un diccionario de Provincia => Alicuota (double), leyendo desde la base SQL.
        /// </summary>
        public static Dictionary<string, double> ObtenerAlicuotasDesdeBD()
        {
            var dict = new Dictionary<string, double>(StringComparer.InvariantCultureIgnoreCase);

            using (var conexion = Automatizacion.Data.ConexionBD.ObtenerConexion())
            {
                var cmd = new SqlCommand("SELECT Provincia, Alicuota FROM zocoweb.dbo.Lista_IIBBProvincia", conexion);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        var provincia = reader["Provincia"]?.ToString()?.Trim();
                        var alicuotaStr = reader["Alicuota"]?.ToString()?.Replace("%", "").Replace(",", ".").Trim();
                        double alicuota = 0;
                        double.TryParse(alicuotaStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out alicuota);

                        if (!string.IsNullOrWhiteSpace(provincia) && !dict.ContainsKey(provincia))
                            dict.Add(provincia, alicuota);
                    }
                }
            }

            return dict;
        }
    }
}
