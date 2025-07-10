using System.Data;
using System.Data.SqlClient;
using Automatizacion.Data;
using Automatizacion.Core.Formularios.Interfaz;

namespace Automatizacion.Core.Formularios.Servicios
{
    public class ProvinciaService : IProvinciaService
    {
        public DataTable ObtenerProvincias()
        {
            using (var conexion = ConexionBD.ObtenerConexion())
            {
                var adaptador = new SqlDataAdapter(
                    @"SELECT TOP 1000 [id], [Provincia], [Alicuota]
                      FROM [zocoweb].[dbo].[Lista_IIBBProvincia]", conexion);

                var tabla = new DataTable();
                adaptador.Fill(tabla);
                return tabla;
            }
        }

        public void GuardarProvincias(DataTable tabla)
        {
            using (var conexion = ConexionBD.ObtenerConexion())
            {
                var adaptador = new SqlDataAdapter(
                    @"SELECT TOP 1000 [id], [Provincia], [Alicuota]
                      FROM [zocoweb].[dbo].[Lista_IIBBProvincia]", conexion);

                var builder = new SqlCommandBuilder(adaptador);
                adaptador.Update(tabla);
            }
        }
    }
}
