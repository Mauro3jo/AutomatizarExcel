using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public static class VisaDebitoProcessor
    {
        public static double Procesar(string rutaArchivo, string nombreHoja, ProgressBar barra, out int filasSumadas)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            double total = 0;
            filasSumadas = 0;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int contador = 0;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celdaH = worksheet.Cells[i, 8] as Excel.Range;
                    string valorH = Normalizar(celdaH?.Value2);

                    if (!string.IsNullOrWhiteSpace(valorH) &&
                        double.TryParse(valorH, NumberStyles.Any, CultureInfo.InvariantCulture, out double bruto))
                    {
                        total += bruto;
                        filasSumadas++;
                    }

                    if (barra != null)
                    {
                        contador++;
                        barra.Invoke((MethodInvoker)(() =>
                        {
                            barra.Value = Math.Min(100, (int)(contador / (float)(lastRow - 1) * 100));
                        }));
                    }
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando Visa Débito:\n" + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

                if (barra != null)
                {
                    barra.Invoke((MethodInvoker)(() => barra.Value = 0));
                }
            }

            return total;
        }



        private static string Normalizar(object valor)
        {
            return Convert.ToString(valor)
                ?.Replace("$", "")
                .Replace(".", "")
                .Replace(",", ".")
                .Trim();
        }
    }
}
