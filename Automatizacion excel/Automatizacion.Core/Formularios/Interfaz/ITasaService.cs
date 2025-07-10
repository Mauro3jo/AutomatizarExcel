using System.Data;

namespace Automatizacion.Core.Formularios.Interfaz
{
    public interface ITasaService
    {
        DataTable ObtenerTasas();
        void GuardarTasas(DataTable tabla);
    }
}
