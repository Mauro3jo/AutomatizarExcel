using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel
{
    public static class CabalProcessor
    {
        public static double Procesar(string rutaArchivo, string nombreHoja)
        {
            double totalBruto = 0;
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celda = worksheet.Cells[i, 6] as Excel.Range; // Columna F = 6
                    string texto = Convert.ToString(celda?.Value2)
                        ?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim();

                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double valor))
                    {
                        totalBruto += valor;
                    }
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando CABAL:\n\n" + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return totalBruto;
        }
    }
}
