using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public static class Amex_2Processor
    {
        public static DataTable ObtenerFilasCandidatas(string rutaArchivo, string nombreHoja, ProgressBar barra = null)
        {
            var dt = new DataTable();
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int lastCol = 45; // Fijo: columna AS

                dt.Columns.Add("FilaExcel", typeof(int));
                for (int col = 1; col <= lastCol; col++)
                {
                    var encabezado = Convert.ToString((worksheet.Cells[1, col] as Excel.Range)?.Value2)?.Trim();
                    dt.Columns.Add(string.IsNullOrWhiteSpace(encabezado) ? $"Col {col}" : encabezado);
                }

                // Solo mostrar candidatas desde fila 2
                for (int i = 2; i <= lastRow; i++)
                {
                    var celdaA = worksheet.Cells[i, 1] as Excel.Range;
                    string valorColA = Convert.ToString(celdaA?.Value2)?.Trim();

                    if (string.IsNullOrWhiteSpace(valorColA))
                    {
                        var fila = dt.NewRow();
                        fila["FilaExcel"] = i;

                        for (int col = 1; col <= lastCol; col++)
                        {
                            var celda = worksheet.Cells[i, col] as Excel.Range;
                            fila[col] = Convert.ToString(celda?.Value2)?.Trim();
                        }

                        dt.Rows.Add(fila);
                    }

                    // Actualizar barra si existe
                    if (barra != null)
                    {
                        int progreso = (int)((i - 1) / (float)(lastRow - 1) * 100);
                        barra.Invoke((MethodInvoker)(() => barra.Value = progreso));
                    }
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error leyendo filas:\n" + ex.Message);
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

            return dt;
        }


        public static double Procesar(string rutaArchivo, string nombreHoja, List<int> filasAEliminar, ProgressBar barra, out int filasContadas)
        {
            filasContadas = 0;
            double totalBruto = 0;
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                // Borrar filas (orden descendente)
                filasAEliminar.Sort();
                filasAEliminar.Reverse();

                int total = filasAEliminar.Count;
                int contador = 0;

                foreach (int fila in filasAEliminar)
                {
                    worksheet.Rows[fila].Delete();
                    contador++;

                    if (barra != null)
                    {
                        barra.Invoke((MethodInvoker)(() =>
                        {
                            barra.Value = (int)(contador / (float)Math.Max(1, total) * 100);
                        }));
                    }
                }

                // Recalcular total bruto: columna AC = col 29
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int contadorSuma = 0;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celda = worksheet.Cells[i, 29] as Excel.Range;
                    string texto = Convert.ToString(celda?.Value2)
                        ?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim();

                    if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double valor))
                    {
                        totalBruto += valor;
                        filasContadas++; // 👈 contar filas válidas
                    }

                    if (barra != null)
                    {
                        contadorSuma++;
                        int progreso = (int)(contadorSuma / (float)(lastRow - 1) * 100);
                        barra.Invoke((MethodInvoker)(() => barra.Value = progreso));
                    }
                }

                workbook.Save();
                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando:\n" + ex.Message);
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
