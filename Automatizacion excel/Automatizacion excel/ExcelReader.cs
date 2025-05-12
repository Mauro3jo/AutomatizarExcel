using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;

namespace Automatizacion_excel
{
    public static class ExcelReader
    {
        public static List<string> ObtenerNombresHojas(string rutaArchivo)
        {
            var hojas = new List<string>();

           
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(rutaArchivo, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet();
                foreach (DataTable table in result.Tables)
                {
                    hojas.Add(table.TableName);
                }
            }

            return hojas;
        }
    }
}
