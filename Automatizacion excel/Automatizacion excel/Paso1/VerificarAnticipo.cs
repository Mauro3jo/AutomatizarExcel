using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public static class VerificarAnticipo
    {
        /// <summary>
        /// Devuelve una lista con los números de fila (desde 2) donde la columna O (anticipo) está vacía pero hay otros datos.
        /// </summary>
        public static List<int> FilasSinAnticipo(string rutaExcel, string hoja)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            List<int> filasVacias = new List<int>();
            try
            {
                var workbook = excelApp.Workbooks.Open(rutaExcel);
                var worksheet = workbook.Sheets[hoja] as Excel.Worksheet;
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int i = 2; i <= lastRow; i++)
                {
                    var celdaO = worksheet.Cells[i, 15] as Excel.Range;
                    string valorO = Convert.ToString(celdaO?.Value2)?.Trim();

                    bool filaConDatos = false;
                    for (int j = 1; j <= 20; j++)
                    {
                        if (j == 15) continue;
                        var celda = worksheet.Cells[i, j] as Excel.Range;
                        string val = Convert.ToString(celda?.Value2)?.Trim();
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            filaConDatos = true;
                            break;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(valorO) && filaConDatos)
                        filasVacias.Add(i);
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error verificando hoja '{hoja}': {ex.Message}", ex);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return filasVacias;
        }

        /// <summary>
        /// Procesa anticipos según tus reglas: guarda únicos en BD (redondeados a 4 decimales), 
        /// pisa columna O en "Plan cuota" y en "Venta ctdo" con K vacía, usando el mayor anticipo.
        /// Si O tiene fórmula, la borra solo en las filas que modifica.
        /// Siempre pone formato porcentaje a la celda modificada (como si fuera "Pegar valores").
        /// </summary>
        public static void ProcesarYActualizarAnticipos(string rutaExcel, string hoja, System.Windows.Forms.Form formularioPrincipal = null)
        {
            var anticiposVenta = new List<double>();
            var filasPlanCuota = new List<int>();
            var filasVentaCtdoSinK = new List<int>();
            bool fila2Coincide = false;

            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                workbook = excelApp.Workbooks.Open(rutaExcel);
                worksheet = workbook.Sheets[hoja] as Excel.Worksheet;
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int i = 2; i <= lastRow; i++)
                {
                    string tipoVenta = Convert.ToString((worksheet.Cells[i, 1] as Excel.Range)?.Value2)?.Trim();
                    string valorAnticipoStr = Convert.ToString((worksheet.Cells[i, 15] as Excel.Range)?.Value2)?.Trim();
                    string valorK = Convert.ToString((worksheet.Cells[i, 11] as Excel.Range)?.Value2)?.Trim();

                    if (tipoVenta != null)
                    {
                        if (tipoVenta.Equals("Venta ctdo", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!string.IsNullOrWhiteSpace(valorK)
                                && !string.IsNullOrWhiteSpace(valorAnticipoStr)
                                && double.TryParse(valorAnticipoStr, out double valorAnticipo))
                            {
                                anticiposVenta.Add(valorAnticipo);
                            }
                            else if (string.IsNullOrWhiteSpace(valorK))
                            {
                                if (i == 2) fila2Coincide = true; // MARCAR QUE LA FILA 2 NO SE MODIFICARÁ
                                else filasVentaCtdoSinK.Add(i);
                            }
                        }
                        else if (tipoVenta.Equals("Plan cuota", StringComparison.OrdinalIgnoreCase))
                        {
                            if (i == 2) fila2Coincide = true; // MARCAR QUE LA FILA 2 NO SE MODIFICARÁ
                            else filasPlanCuota.Add(i);
                        }
                    }
                }

                // Guardar anticipos únicos en la base de datos (redondeados a 4 decimales)
                var anticiposUnicosRedondeados = anticiposVenta
                    .Select(a => Math.Round(a, 4, MidpointRounding.AwayFromZero))
                    .Distinct()
                    .ToList();

                GuardarAnticiposEnBaseDeDatos(anticiposUnicosRedondeados, hoja); // <-- Ahora le pasamos la tarjeta

                // Obtener el mayor anticipo y pisar la columna O de "Plan cuota" y de "Venta ctdo" con K vacía
                if (anticiposVenta.Any() && (filasPlanCuota.Any() || filasVentaCtdoSinK.Any()))
                {
                    double mayorAnticipo = anticiposVenta.Max();

                    foreach (var fila in filasPlanCuota.Concat(filasVentaCtdoSinK))
                    {
                        var celdaO = worksheet.Cells[fila, 15] as Excel.Range;

                        // Borra fórmula si hay, para que quede sólo valor (efecto igual a pegar valores)
                        if (celdaO.HasFormula)
                            celdaO.Clear();

                        // Asigna valor y SIEMPRE formato porcentaje, como "pegar valores" + "formato %"
                        celdaO.Value2 = mayorAnticipo;
                        celdaO.NumberFormat = "#,##0.00%";
                    }
                    worksheet.Columns[15].AutoFit();
                }

                workbook.Save();

                // AVISAR si la fila 2 cumplía los criterios pero NO se tocó
                if (fila2Coincide)
                {
                    string nombreVisible = hoja == "Visa" ? "Visa Crédito"
                        : hoja == "Mastercard" ? "Mastercard Crédito"
                        : hoja == "ARGENCARD" ? "ArgenCard"
                        : hoja;

                    string mensaje = $"En la hoja \"{nombreVisible}\", la fila 2 cumple las condiciones de reemplazo pero NO se modifica para preservar la fórmula original.";
                    if (formularioPrincipal != null)
                        System.Windows.Forms.MessageBox.Show(formularioPrincipal, mensaje, "Atención", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    else
                        System.Windows.Forms.MessageBox.Show(mensaje, "Atención", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        /// <summary>
        /// Guarda anticipos en la base de datos con la fecha/hora actual de Argentina **y la tarjeta**.
        /// </summary>
        private static void GuardarAnticiposEnBaseDeDatos(List<double> anticipos, string tarjeta)
        {
            if (anticipos == null || anticipos.Count == 0)
                return;

            using (var conn = Automatizacion.Data.ConexionBD.ObtenerConexion())
            {
                foreach (var anticipo in anticipos)
                {
                    using (var cmd = new SqlCommand("INSERT INTO anticipo (anticipo, fecha, Tarjeta) VALUES (@anticipo, @fecha, @tarjeta)", conn))
                    {
                        cmd.Parameters.AddWithValue("@anticipo", anticipo);
                        var fechaArgentina = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.UtcNow, "Argentina Standard Time");
                        cmd.Parameters.AddWithValue("@fecha", fechaArgentina);
                        cmd.Parameters.AddWithValue("@tarjeta", tarjeta); // 👈 Nuevo campo
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}
