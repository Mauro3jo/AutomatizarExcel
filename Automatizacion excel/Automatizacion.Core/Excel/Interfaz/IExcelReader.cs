using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatizacion.Core.Excel.Interfaz
{
    public interface IExcelReader
    {
        List<string> ObtenerNombresHojas(string rutaArchivo);
    }

}
