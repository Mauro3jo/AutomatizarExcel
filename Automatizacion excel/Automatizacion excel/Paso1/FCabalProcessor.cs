using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public static class CabalProcessor
    {
        public static double Procesar(string rutaArchivo, string nombreHoja, ProgressBar barra, out int filasContadas)
        {
            filasContadas = 0;
            double totalBruto = 0;
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int contador = 0;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celda = worksheet.Cells[i, 6] as Excel.Range; // Columna F = 6
                    string texto = Convert.ToString(celda?.Value2)
                        ?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim();

                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double valor))
                    {
                        totalBruto += valor;
                        filasContadas++; // 👈 contar fila válida
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
                MessageBox.Show("Error procesando CABAL:\n\n" + ex.Message);
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

            return totalBruto;
        }


    }
}
