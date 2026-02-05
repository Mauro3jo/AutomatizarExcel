using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using Automatizacion.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso4
{
    /// <summary>
    /// Control de tasas financieras contra Lista_Cuota
    /// para el REPORTE DIARIO (Paso 4).
    /// TODO está contenido acá.
    /// </summary>
    public static class ControlTasasReporteDiario
    {
        // ============================================================
        // 1️⃣ OBTENER TASAS DESDE BD
        // ============================================================
        public static Dictionary<(string tarjeta, int cuota), double> ObtenerTasas()
        {
            var tasas = new Dictionary<(string tarjeta, int cuota), double>();

            using (var con = ConexionBD.ObtenerConexion())
            {
                var cmd = new SqlCommand(@"
                    SELECT
                        Cuota,
                        Costo_Visa_Credito,
                        Costo_American_Express_Credito,
                        Costo_Mastercard_Credito,
                        Costo_Argencard_Credito,
                        Costo_Cabal_Credito,
                        Costo_Naranja_Credito
                    FROM zocoweb.dbo.Lista_Cuota
                    WHERE ISNUMERIC(Cuota) = 1
                ", con);

                using (var rd = cmd.ExecuteReader())
                {
                    while (rd.Read())
                    {
                        if (!int.TryParse(rd["Cuota"]?.ToString(), out int cuota))
                            continue;

                        CargarTasa(tasas, "VISA", cuota, rd["Costo_Visa_Credito"]);
                        CargarTasa(tasas, "AMERICAN EXPRESS", cuota, rd["Costo_American_Express_Credito"]);
                        CargarTasa(tasas, "MASTERCARD", cuota, rd["Costo_Mastercard_Credito"]);
                        CargarTasa(tasas, "ARGENCARD", cuota, rd["Costo_Argencard_Credito"]);
                        CargarTasa(tasas, "CABAL", cuota, rd["Costo_Cabal_Credito"]);
                        CargarTasa(tasas, "NARANJA", cuota, rd["Costo_Naranja_Credito"]);
                    }
                }
            }

            return tasas;
        }

        private static void CargarTasa(
            Dictionary<(string tarjeta, int cuota), double> dict,
            string tarjeta,
            int cuota,
            object valorBD
        )
        {
            if (valorBD == null) return;

            string str = valorBD.ToString()
                .Replace("%", "")
                .Replace(",", ".")
                .Trim();

            if (string.IsNullOrWhiteSpace(str)) return;
            if (str.Equals("Revisar", StringComparison.OrdinalIgnoreCase)) return;

            if (double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out double tasa))
            {
                dict[(tarjeta.ToUpper(), cuota)] = tasa / 100.0;
            }
        }

        // ============================================================
        // 2️⃣ VERIFICAR TASAS CONTRA REPORTE DIARIO2
        // ============================================================
        public static List<int> VerificarTasas(
            string rutaExcel,
            Dictionary<(string tarjeta, int cuota), double> tasas,
            Action<string, int> reportarProgreso = null
        )
        {
            var filasConError = new List<int>();

            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application { DisplayAlerts = false };
                wb = excelApp.Workbooks.Open(rutaExcel);

                // 🔥 NOMBRE CORRECTO
                ws = wb.Sheets["Reporte Diario2"] as Excel.Worksheet;
                if (ws == null)
                    return filasConError;

                int lastRow;
                try
                {
                    lastRow = ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                }
                catch
                {
                    return filasConError; // hoja vacía
                }

                for (int i = 2; i <= lastRow; i++)
                {
                    try
                    {
                        // Q = 17 → Cuota
                        if (!int.TryParse(
                            Convert.ToString((ws.Cells[i, 17] as Excel.Range)?.Value2),
                            out int cuota))
                            continue;

                        // W = 23 → Tarjeta
                        string tarjeta = Convert.ToString((ws.Cells[i, 23] as Excel.Range)?.Value2)
                            ?.Trim()
                            ?.ToUpper();

                        if (string.IsNullOrEmpty(tarjeta))
                            continue;

                        // X = 24 → Costo financiero
                        string strCosto = Convert.ToString((ws.Cells[i, 24] as Excel.Range)?.Value2)
                            ?.Replace("%", "")
                            ?.Replace(",", ".");

                        if (!double.TryParse(strCosto, NumberStyles.Any, CultureInfo.InvariantCulture, out double costoExcel))
                            continue;

                        costoExcel /= 100.0;

                        if (!tasas.TryGetValue((tarjeta, cuota), out double tasaPermitida))
                            continue;

                        if (costoExcel > tasaPermitida)
                            filasConError.Add(i);

                        reportarProgreso?.Invoke(
                            $"📊 Controlando tasas fila {i}",
                            (int)((i / (double)lastRow) * 100)
                        );
                    }
                    catch
                    {
                        // error puntual → no rompe nada
                        continue;
                    }
                }
            }
            finally
            {
                try
                {
                    if (ws != null) Marshal.ReleaseComObject(ws);
                    if (wb != null)
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(wb);
                    }
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }

            return filasConError;
        }
    }
}
