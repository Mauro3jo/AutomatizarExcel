using System.Data;
using System.Data.SqlClient;
using Automatizacion.Data;
using Automatizacion.Core.Formularios.Interfaz;

namespace Automatizacion.Core.Formularios.Servicios
{
    public class TasaService : ITasaService
    {
        public DataTable ObtenerTasas()
        {
            using (var conexion = ConexionBD.ObtenerConexion())
            {
                var adaptador = new SqlDataAdapter(
                    @"SELECT TOP 1000 [Id], [Cuota], [Codigo_Posnet],
                            [Costo_Visa_Credito], [Costo_American_Express_Credito],
                            [Costo_Mastercard_Credito], [Costo_Argencard_Credito], 
                            [Costo_Cabal_Credito], [Costo_Naranja_Credito], 
                            [Comision], [IVA], [Comision_mas_IVA]
                      FROM [zocoweb].[dbo].[Lista_Cuota]",
                    conexion);

                var tabla = new DataTable();
                adaptador.Fill(tabla);
                return tabla;
            }
        }

        public void GuardarTasas(DataTable tabla)
        {
            using (var conexion = ConexionBD.ObtenerConexion())
            {
                var adaptador = new SqlDataAdapter(
                    @"SELECT TOP 1000 [Id], [Cuota], [Codigo_Posnet],
                            [Costo_Visa_Credito], [Costo_American_Express_Credito],
                            [Costo_Mastercard_Credito], [Costo_Argencard_Credito], 
                            [Costo_Cabal_Credito], [Costo_Naranja_Credito], 
                            [Comision], [IVA], [Comision_mas_IVA]
                      FROM [zocoweb].[dbo].[Lista_Cuota]",
                    conexion);

                var builder = new SqlCommandBuilder(adaptador);
                adaptador.Update(tabla);
            }
        }
    }
}
