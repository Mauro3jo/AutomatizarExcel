using System;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Automatizacion.Data 
{
    public static class ConexionBD
    {
        private static IConfigurationRoot? _configuration;

        private static IConfigurationRoot Configuration
        {
            get
            {
                if (_configuration == null)
                {
                    _configuration = new ConfigurationBuilder()
                        .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                        .Build();
                }
                return _configuration;
            }
        }

        public static SqlConnection ObtenerConexion()
        {
            string cadenaConexion = Configuration.GetConnectionString("MiConexion");

            if (string.IsNullOrWhiteSpace(cadenaConexion))
                throw new InvalidOperationException("La cadena de conexión 'MiConexion' no fue encontrada o está vacía. Revisá el appsettings.json y la copia en el directorio de salida.");

            var conexion = new SqlConnection(cadenaConexion);
            conexion.Open();
            return conexion;
        }
    }
}
