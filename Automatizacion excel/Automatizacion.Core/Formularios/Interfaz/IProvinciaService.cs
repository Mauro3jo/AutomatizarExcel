using System.Data;

namespace Automatizacion.Core.Formularios.Interfaz
{
    public interface IProvinciaService
    {
        DataTable ObtenerProvincias();
        void GuardarProvincias(DataTable tabla);
    }
}
