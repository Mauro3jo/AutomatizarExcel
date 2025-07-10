using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using Automatizacion.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1
{
    public static class ControlTasas
    {
        /// <summary>
        /// Devuelve las tasas por cuota desde la base, para una tarjeta
        /// tarjeta puede ser: "Visa", "Mastercard", "ARGENCARD"
        /// </summary>
        public static Dictionary<int, double> ObtenerTasasDesdeBD(string tarjeta)
        {
            string campo = tarjeta.ToUpper() switch
            {
                "VISA" => "Costo_Visa_Credito",
                "MASTERCARD" => "Costo_Mastercard_Credito",
                "ARGENCARD" => "Costo_Argencard_Credito",
                _ => throw new Exception("Tarjeta no soportada")
            };

            var tasas = new Dictionary<int, double>();

            using (var conexion = ConexionBD.ObtenerConexion())
            {
                var cmd = new SqlCommand($@"
                    SELECT [Cuota], [{campo}]
                    FROM [zocoweb].[dbo].[Lista_Cuota]
                    WHERE ISNUMERIC([Cuota]) = 1", conexion);

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        if (int.TryParse(reader["Cuota"].ToString(), out int cuota))
                        {
                            var strTasa = reader[campo]?.ToString()?.Replace("%", "").Replace(",", ".").Trim();

                            // Solo agrego si NO es vacío, ni "Revisar"
                            if (!string.IsNullOrWhiteSpace(strTasa) &&
                                !strTasa.Equals("Revisar", StringComparison.OrdinalIgnoreCase) &&
                                double.TryParse(strTasa, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double tasaVal))
                            {
                                tasas[cuota] = tasaVal / 100.0; // paso a decimal (por ejemplo 4.69% => 0.0469)
                            }
                        }
                    }
                }
            }
            return tasas;
        }

        /// <summary>
        /// Devuelve la lista de filas (índices Excel) donde el ratio supera el máximo permitido.
        /// El ratio debe ser MENOR al total permitido (anticipo + iva + extra + tasa).
        /// Si es MAYOR O IGUAL, es error.
        /// </summary>
        public static List<int> VerificarExcesos(
            string rutaExcel,
            string hoja,
            Dictionary<int, double> tasasPorCuota // cuota -> tasa en decimal (ej: 0.0469 para 4,69%)
        )
        {
            var filasConExceso = new List<int>();

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
                    string tipoOperacion = Convert.ToString((worksheet.Cells[i, 1] as Excel.Range)?.Value2)?.Trim();
                    if (!string.Equals(tipoOperacion, "Plan cuota", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Columna E: cuota
                    string strCuota = Convert.ToString((worksheet.Cells[i, 5] as Excel.Range)?.Value2)?.Trim();
                    if (!int.TryParse(strCuota, out int cuota))
                        continue;

                    // Ajuste para cuotas 13->3 y 16->6
                    if (cuota == 13) cuota = 3;
                    if (cuota == 16) cuota = 6;

                    // Tasa para la cuota
                    if (!tasasPorCuota.TryGetValue(cuota, out double tasaCuota))
                        continue;

                    // Col H (8): monto 1
                    string strH = Convert.ToString((worksheet.Cells[i, 8] as Excel.Range)?.Value2)?.Trim();
                    // Col K (11): monto base
                    string strK = Convert.ToString((worksheet.Cells[i, 11] as Excel.Range)?.Value2)?.Trim();
                    // Col O (15): anticipo
                    string strO = Convert.ToString((worksheet.Cells[i, 15] as Excel.Range)?.Value2)?.Trim();

                    if (!double.TryParse(strH, out double montoH) ||
        !double.TryParse(strK, out double montoK) ||
        montoH == 0)
                        continue;

                    // Porcentaje de descuento respecto al bruto
                    double porcentajeDescuento = montoK / montoH;

                    // Anticipo
                    if (!double.TryParse(strO, out double anticipo))
                        anticipo = 0;

                    // IVA sobre anticipo (21%)
                    double iva = anticipo * 0.21;
                    double extra = 0.005; // 0,5%

                    // SUMA TOTAL: anticipo + iva + extra + tasa de cuota
                    double totalComparar = anticipo + iva + extra + tasaCuota;

                    // Si el porcentaje de descuento es MAYOR que lo permitido, está mal
                    if (porcentajeDescuento > totalComparar)
                        filasConExceso.Add(i);
 


                }

                workbook.Close(false);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return filasConExceso;
        }
    }
}
