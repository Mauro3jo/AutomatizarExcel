using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1QR
{
    /// <summary>
    /// Servicio para ejecutar macros de Excel sobre un archivo dado.
    /// </summary>
    internal static class MacroQR
    {
        /// <summary>
        /// Ejecuta una macro por nombre en el archivo Excel especificado.
        /// </summary>
        /// <param name="rutaExcel">Ruta completa del archivo Excel (.xlsm).</param>
        /// <param name="nombreMacro">Nombre exacto de la macro a ejecutar.</param>
        public static void EjecutarMacro(string rutaExcel, string nombreMacro)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbook wb = null;
            try
            {
                wb = excelApp.Workbooks.Open(rutaExcel);
                excelApp.Run(nombreMacro);
                wb.Save();
                MessageBox.Show($"Macro '{nombreMacro}' ejecutada correctamente.", "Macro ejecutada", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al ejecutar la macro '{nombreMacro}':\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (wb != null) { wb.Close(); Marshal.ReleaseComObject(wb); }
                excelApp.Quit(); Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
