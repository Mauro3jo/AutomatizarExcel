﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel
{
    public static class ArgencardProcessor
    {
        public static DataTable ObtenerFilasAfectadas(string rutaArchivo, string nombreHoja)
        {
            var dt = new DataTable();
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int lastCol = 45;

                dt.Columns.Add("FilaExcel", typeof(int));
                for (int col = 1; col <= lastCol; col++)
                {
                    var encabezado = Convert.ToString((worksheet.Cells[1, col] as Excel.Range)?.Value2)?.Trim();
                    dt.Columns.Add(string.IsNullOrWhiteSpace(encabezado) ? $"Col {col}" : encabezado);
                }

                for (int i = 2; i <= lastRow; i++)
                {
                    var celdaE = worksheet.Cells[i, 5] as Excel.Range;
                    string valorE = Convert.ToString(celdaE?.Value2)?.Trim();

                    if (string.IsNullOrWhiteSpace(valorE)) continue;

                    var fila = dt.NewRow();
                    fila["FilaExcel"] = i;
                    for (int col = 1; col <= lastCol; col++)
                    {
                        var celda = worksheet.Cells[i, col] as Excel.Range;
                        fila[col] = Convert.ToString(celda?.Value2)?.Trim();
                    }
                    dt.Rows.Add(fila);
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error obteniendo vista previa:\n" + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return dt;
        }

        public static double Procesar(string rutaArchivo, string nombreHoja, List<int> filasSeleccionadas, ProgressBar barra = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            double total = 0;

            try
            {
                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var worksheet = workbook.Sheets[nombreHoja] as Excel.Worksheet;

                filasSeleccionadas.Sort();
                filasSeleccionadas.Reverse();

                var filasValidas = new List<int>();

                foreach (int fila in filasSeleccionadas)
                {
                    var celdaE = worksheet.Cells[fila, 5] as Excel.Range;
                    string valorE = Convert.ToString(celdaE?.Value2)?.Trim();

                    if (string.IsNullOrWhiteSpace(valorE)) continue;

                    if (valorE.Contains("/"))
                    {
                        if (!valorE.StartsWith("01/"))
                        {
                            worksheet.Rows[fila].Delete();
                            continue;
                        }
                    }

                    filasValidas.Add(fila);
                }

                filasValidas.Sort();
                int totalOperaciones = filasValidas.Count;
                int contador = 0;

                foreach (int fila in filasValidas)
                {
                    var celdaE = worksheet.Cells[fila, 5] as Excel.Range;
                    string valorE = Convert.ToString(celdaE?.Value2)?.Trim();
                    int cuotas = 1;
                    bool debeMultiplicar = false;

                    if (valorE.Contains("/"))
                    {
                        var partes = valorE.Split('/');
                        if (!int.TryParse(partes[1], out cuotas)) cuotas = 1;
                        debeMultiplicar = true;
                    }
                    else
                    {
                        if (!int.TryParse(valorE, out cuotas)) cuotas = 1;
                        debeMultiplicar = false;
                    }

                    string nuevoTextoE = cuotas == 3 ? "13" :
                                         cuotas == 6 ? "16" :
                                         cuotas.ToString();
                    worksheet.Cells[fila, 5].Value2 = nuevoTextoE;

                    var celdaH = worksheet.Cells[fila, 8] as Excel.Range;
                    string textoH = Normalizar(celdaH?.Value2);
                    if (!string.IsNullOrWhiteSpace(textoH) &&
                        double.TryParse(textoH, NumberStyles.Any, CultureInfo.InvariantCulture, out double valorH))
                    {
                        double nuevoValorH = debeMultiplicar ? valorH * cuotas : valorH;
                        worksheet.Cells[fila, 8].Value2 = nuevoValorH;
                    }

                    var celdaJ = worksheet.Cells[fila, 10] as Excel.Range;
                    string textoJ = Normalizar(celdaJ?.Value2);
                    if (!string.IsNullOrWhiteSpace(textoJ) &&
                        double.TryParse(textoJ, NumberStyles.Any, CultureInfo.InvariantCulture, out double valorJ))
                    {
                        double nuevoValorJ = debeMultiplicar ? valorJ * cuotas : valorJ;
                        worksheet.Cells[fila, 10].Value2 = nuevoValorJ;
                    }

                    if (barra != null)
                    {
                        contador++;
                        barra.Invoke((MethodInvoker)(() =>
                        {
                            barra.Value = (int)((contador / (float)totalOperaciones) * 100);
                        }));
                    }
                }

                // ✅ Sumar toda la columna H final, sin importar cuotas
                int lastRowFinal = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                for (int i = 2; i <= lastRowFinal; i++)
                {
                    var celdaH = worksheet.Cells[i, 8] as Excel.Range;
                    string brutoTxt = Normalizar(celdaH?.Value2);
                    if (!string.IsNullOrWhiteSpace(brutoTxt) &&
                        double.TryParse(brutoTxt, NumberStyles.Any, CultureInfo.InvariantCulture, out double bruto))
                    {
                        total += bruto;
                    }
                }

                workbook.Save();
                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando ARGENCARD:\n" + ex.Message);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
                if (barra != null)
                    barra.Invoke((MethodInvoker)(() => barra.Value = 0));
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
