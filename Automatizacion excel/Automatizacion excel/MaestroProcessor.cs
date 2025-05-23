﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel
{
    public static class MaestroProcessor
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
                    var celdaVentas = worksheet.Cells[i, 8] as Excel.Range;
                    var textoVentas = Convert.ToString(celdaVentas?.Value2)?.Trim();

                    var celdaC = worksheet.Cells[i, 3] as Excel.Range;
                    string textoC = Convert.ToString(celdaC?.Value2)?.Trim();

                    if (string.IsNullOrWhiteSpace(textoVentas) &&
                        !string.IsNullOrWhiteSpace(textoC) &&
                        textoC.Split(' ').Length == 3)
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
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error obteniendo filas MAESTRO:\n" + ex.Message);
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
                int totalOperaciones = filasSeleccionadas.Count;
                int contador = 0;

                foreach (int fila in filasSeleccionadas)
                {
                    var celdaC = worksheet.Cells[fila, 3] as Excel.Range;
                    string textoC = Convert.ToString(celdaC?.Value2)?.Trim();

                    if (!string.IsNullOrWhiteSpace(textoC) && textoC.Split(' ').Length == 3)
                    {
                        // Primero: Mover F → H, E → G, D → F
                        worksheet.Cells[fila, 8].Value2 = worksheet.Cells[fila, 6]?.Value2; // F → H
                        worksheet.Cells[fila, 7].Value2 = worksheet.Cells[fila, 5]?.Value2; // E → G
                        worksheet.Cells[fila, 6].Value2 = worksheet.Cells[fila, 4]?.Value2; // D → F

                        // Luego, dividir C en 3 partes
                        var partes = textoC.Split(' ');
                        worksheet.Cells[fila, 3].Value2 = partes[0]; // Term
                        worksheet.Cells[fila, 4].Value2 = partes[1]; // Lote
                        worksheet.Cells[fila, 5].Value2 = partes[2]; // Cupon
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

                // Sumar toda la columna G (Ventas)
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                for (int i = 2; i <= lastRow; i++)
                {
                    var celdaG = worksheet.Cells[i, 7] as Excel.Range;
                    string valorG = Convert.ToString(celdaG?.Value2)?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim();

                    if (double.TryParse(valorG, NumberStyles.Any, CultureInfo.InvariantCulture, out double bruto))
                        total += bruto;
                }

                workbook.Save();
                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error procesando MAESTRO:\n" + ex.Message);
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
    }
}
