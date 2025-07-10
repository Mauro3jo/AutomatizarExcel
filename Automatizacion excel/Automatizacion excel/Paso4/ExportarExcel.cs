using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso4
{
    internal class ExportarExcel
    {
        public void ExportarResumenComisiones(string archivoOrigen, string archivoDestino)
        {
            Excel.Application excelApp = null;
            Excel.Workbook wbOrigen = null;
            Excel.Workbook wbDestino = null;

            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                wbOrigen = excelApp.Workbooks.Open(archivoOrigen, ReadOnly: true);
                Excel.Worksheet hoja = wbOrigen.Sheets["Reporte Diario2"];

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // Columnas útiles (índices base 1)
                int COL_BRUTO = 8;        // H
                int COL_CUOTAS = 17;      // Q
                int COL_PORC_CF = 24;     // X  (% Costo Financiero, para mostrar el "real")
                int COL_COSTO_FINAN = 25; // Y  (Costo Financiero $)
                int COL_ANTICIPO = 27;    // AA (Anticipo)
                int COL_COMISION = 29;    // AC
                int COL_IVA = 30;         // AD

                // Tabla débito/crédito
                double sumaBrutoDebito = 0, sumaComisionDebito = 0, sumaIvaDebito = 0;
                double sumaBrutoCredito = 0, sumaComisionCredito = 0, sumaIvaCredito = 0;

                // Para cuotas: incluyendo SumaAnticipo
                var datosCuotas = new Dictionary<int, (double Bruto, double CostoFinan, double? PorcReal, double SumaAnticipo)>();

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    double bruto = LeerDouble(hoja.Cells[fila, COL_BRUTO]);
                    int cuotas = (int)LeerDouble(hoja.Cells[fila, COL_CUOTAS]);
                    double comision = LeerDouble(hoja.Cells[fila, COL_COMISION]);
                    double iva = LeerDouble(hoja.Cells[fila, COL_IVA]);
                    double costoFinan = LeerDouble(hoja.Cells[fila, COL_COSTO_FINAN]);
                    double porcReal = LeerPorcentaje(hoja.Cells[fila, COL_PORC_CF]);
                    double anticipo = LeerDouble(hoja.Cells[fila, COL_ANTICIPO]);

                    // Tabla 1: Débito/Crédito
                    if (cuotas == 0)
                    {
                        sumaBrutoDebito += bruto;
                        sumaComisionDebito += comision;
                        sumaIvaDebito += iva;
                    }
                    else
                    {
                        sumaBrutoCredito += bruto;
                        sumaComisionCredito += comision;
                        sumaIvaCredito += iva;
                    }

                    // Tabla 2: por cuotas
                    if (!datosCuotas.ContainsKey(cuotas))
                        datosCuotas[cuotas] = (0, 0, null, 0);

                    var tupla = datosCuotas[cuotas];
                    tupla.Bruto += bruto;
                    tupla.CostoFinan += costoFinan;
                    tupla.SumaAnticipo += anticipo;

                    // % Real (solo el primer valor válido encontrado para cada cuota)
                    if (tupla.PorcReal == null && !string.IsNullOrWhiteSpace(hoja.Cells[fila, COL_PORC_CF].Text))
                        tupla.PorcReal = porcReal;

                    datosCuotas[cuotas] = tupla;
                }

                // Exportar resumen
                wbDestino = excelApp.Workbooks.Add();
                Excel.Worksheet ws = wbDestino.Sheets[1];
                ws.Name = "Resumen";

                int f = 1;

                // --- Tabla Débito/Crédito ---
                ws.Cells[f, 1] = "Tipo";
                ws.Cells[f, 2] = "Suma Bruto";
                ws.Cells[f, 3] = "Suma Comisión";
                ws.Cells[f, 4] = "% Comisión";
                ws.Cells[f, 5] = "Suma IVA";
                ws.Cells[f, 6] = "% IVA";
                ws.Cells[f, 7] = "Comisión + IVA";
                ws.Cells[f, 8] = "% Total sobre Bruto";
                f++;

                // Débito
                ws.Cells[f, 1] = "Débito";
                ws.Cells[f, 2] = sumaBrutoDebito;
                ws.Cells[f, 3] = sumaComisionDebito;
                ws.Cells[f, 4] = sumaBrutoDebito > 0 ? (sumaComisionDebito / sumaBrutoDebito) : 0;
                ws.Cells[f, 5] = sumaIvaDebito;
                ws.Cells[f, 6] = sumaComisionDebito > 0 ? (sumaIvaDebito / sumaComisionDebito) : 0; // %IVA sobre Comisión
                double baseIvaDebito = sumaComisionDebito + sumaIvaDebito;
                ws.Cells[f, 7] = baseIvaDebito;
                ws.Cells[f, 8] = sumaBrutoDebito > 0 ? (baseIvaDebito / sumaBrutoDebito) : 0;
                f++;

                // Crédito
                ws.Cells[f, 1] = "Crédito";
                ws.Cells[f, 2] = sumaBrutoCredito;
                ws.Cells[f, 3] = sumaComisionCredito;
                ws.Cells[f, 4] = sumaBrutoCredito > 0 ? (sumaComisionCredito / sumaBrutoCredito) : 0;
                ws.Cells[f, 5] = sumaIvaCredito;
                ws.Cells[f, 6] = sumaComisionCredito > 0 ? (sumaIvaCredito / sumaComisionCredito) : 0; // %IVA sobre Comisión
                double baseIvaCredito = sumaComisionCredito + sumaIvaCredito;
                ws.Cells[f, 7] = baseIvaCredito;
                ws.Cells[f, 8] = sumaBrutoCredito > 0 ? (baseIvaCredito / sumaBrutoCredito) : 0;
                f += 2;

                // --- Tabla por cuotas ---
                ws.Cells[f, 1] = "Cuotas";
                ws.Cells[f, 2] = "Suma Bruto";
                ws.Cells[f, 3] = "Suma Costo Financiero";
                ws.Cells[f, 4] = "Suma Anticipo";
                ws.Cells[f, 5] = "% Calculado";
                ws.Cells[f, 6] = "% Real (primera fila)";
                ws.Cells[f, 7] = "% Anticipo / Bruto";
                ws.Cells[f, 8] = "% Total Costo Financiero + Anticipo";
                f++;

                int filaInicioCuotas = f;

                foreach (var cuota in datosCuotas.Keys.OrderBy(x => x))
                {
                    var tupla = datosCuotas[cuota];
                    double porcCalculado = tupla.Bruto > 0 ? (tupla.CostoFinan / tupla.Bruto) : 0;
                    double porcReal = tupla.PorcReal ?? 0;
                    double porcAnticipo = tupla.Bruto > 0 ? (tupla.SumaAnticipo / tupla.Bruto) : 0;
                    double porcTotal = porcCalculado + porcAnticipo;

                    ws.Cells[f, 1] = cuota;
                    ws.Cells[f, 2] = tupla.Bruto;
                    ws.Cells[f, 3] = tupla.CostoFinan;
                    ws.Cells[f, 4] = tupla.SumaAnticipo;
                    ws.Cells[f, 5] = porcCalculado;
                    ws.Cells[f, 6] = porcReal;
                    ws.Cells[f, 7] = porcAnticipo;
                    ws.Cells[f, 8] = porcTotal;
                    f++;
                }

                int filaFinCuotas = f - 1;

                // ---- FORMATO MONEDA/PORCENTAJE ----

                // Formato para tabla débito/crédito (robusto, celda por celda)
                for (int row = 2; row <= 3; row++)
                {
                    ws.Cells[row, 2].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 3].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 5].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 7].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 4].NumberFormat = "0.00%";
                    ws.Cells[row, 6].NumberFormat = "0.00%";
                    ws.Cells[row, 8].NumberFormat = "0.00%";
                }

                // Formato para tabla cuotas (solo donde hay datos)
                for (int row = filaInicioCuotas; row <= filaFinCuotas; row++)
                {
                    ws.Cells[row, 2].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 3].NumberFormat = "$ #,##0.00";
                    ws.Cells[row, 4].NumberFormat = "$ #,##0.00";      // Suma Anticipo
                    ws.Cells[row, 5].NumberFormat = "0.00%";
                    ws.Cells[row, 6].NumberFormat = "0.00%";
                    ws.Cells[row, 7].NumberFormat = "0.00%";
                    ws.Cells[row, 8].NumberFormat = "0.00%";           // % Total Costo Financiero + Anticipo
                }

                // Autoajustar ancho columnas
                ws.Columns.AutoFit();

                wbDestino.SaveAs(archivoDestino);
            }
            finally
            {
                wbDestino?.Close(false);
                wbOrigen?.Close(false);
                excelApp?.Quit();
                if (wbDestino != null) Marshal.ReleaseComObject(wbDestino);
                if (wbOrigen != null) Marshal.ReleaseComObject(wbOrigen);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private double LeerDouble(dynamic celda)
        {
            try
            {
                if (celda == null || celda.Value2 == null) return 0;
                string texto = celda.Value2.ToString().Replace("$", "").Replace("%", "").Replace(".", "").Replace(",", ".").Trim();
                double.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double valor);
                return valor;
            }
            catch { return 0; }
        }

        private double LeerPorcentaje(dynamic celda)
        {
            try
            {
                if (celda == null || celda.Value2 == null) return 0;
                string texto = celda.Value2.ToString()
                    .Replace("%", "")
                    .Replace(",", ".")
                    .Trim();

                // Si viene como 0.0469 (decimal)
                if (texto.Contains(".") && texto.Split('.')[1].Length > 2)
                {
                    double.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double valor);
                    if (valor <= 1) return valor;      // ya es decimal
                    else return valor / 100.0;         // venía como 4.69, paso a 0.0469
                }
                else
                {
                    // Si viene como "4" o "4.69"
                    double.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double valor);
                    return valor / 100.0;
                }
            }
            catch { return 0; }
        }
    }
}
 