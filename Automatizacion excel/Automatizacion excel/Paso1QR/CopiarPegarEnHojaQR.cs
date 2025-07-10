using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1QR
{
    internal class CopiarPegarEnHojaQR
    {
        /// <summary>
        /// Copia todas las filas (excepto la cabecera) de la hoja 1 del archivo origen 
        /// y las pega como valores al final de la hoja "QR" del archivo destino.
        /// </summary>
        /// <param name="rutaDestino">Ruta del archivo donde se pega (tiene la hoja QR).</param>
        /// <param name="rutaOrigen">Ruta del archivo desde donde se copia (hoja 1).</param>
        /// <returns>true si todo fue bien, false si hubo error o nada para copiar.</returns>
        public static bool Ejecutar(string rutaDestino, string rutaOrigen)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            Excel.Workbook wbDestino = null;
            Excel.Workbook wbOrigen = null;
            try
            {
                wbDestino = excelApp.Workbooks.Open(rutaDestino);
                wbOrigen = excelApp.Workbooks.Open(rutaOrigen);

                Excel.Worksheet hojaDestino = wbDestino.Sheets["QR"] as Excel.Worksheet;
                Excel.Worksheet hojaOrigen = wbOrigen.Sheets[1] as Excel.Worksheet; // Primera hoja

                // Última fila usada en origen
                int ultimaFilaOrigen = hojaOrigen.Cells[hojaOrigen.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                if (ultimaFilaOrigen < 2)
                {
                    MessageBox.Show("No hay filas para copiar en el archivo origen.", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                // Tomar desde la fila 2 hasta el final, columnas A a V (1 a 22)
                Excel.Range rangoOrigen = hojaOrigen.Range["A2", hojaOrigen.Cells[ultimaFilaOrigen, 22]];

                // Buscar la primera fila vacía en destino
                int ultimaFilaDestino = hojaDestino.Cells[hojaDestino.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row;
                if (ultimaFilaDestino < 1) ultimaFilaDestino = 1;
                int filaInicioPegado = ultimaFilaDestino + 1;

                Excel.Range rangoDestino = hojaDestino.Cells[filaInicioPegado, 1];

                // Pegar solo los valores
                rangoDestino.Resize[rangoOrigen.Rows.Count, rangoOrigen.Columns.Count].Value = rangoOrigen.Value;

                wbDestino.Save();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al copiar/pegar datos QR:\n\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (wbOrigen != null) { wbOrigen.Close(false); Marshal.ReleaseComObject(wbOrigen); }
                if (wbDestino != null) { wbDestino.Close(); Marshal.ReleaseComObject(wbDestino); }
                excelApp.Quit(); Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
