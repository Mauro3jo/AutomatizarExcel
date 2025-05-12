using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel
{
    public static class AmexFiservProcessor
    {
        public static double Procesar(string rutaArchivo, string nombreHoja)
        {
            double total = 0;
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celda = worksheet.Cells[i, 8] as Excel.Range; // Columna H = 8
                    string texto = Convert.ToString(celda?.Value2)
                        ?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim();

                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double valor))
                    {
                        total += valor;
                    }
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando AMEX FISERV:\n\n" + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return total;
        }
    }
}
