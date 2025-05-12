using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel
{
    public class OperacionesPorFechaService
    {
        public void ProcesarOperaciones(string rutaSas, DateTime fechaCorte, string carpetaQuitadas, string carpetaAgregadas)
        {
            string fechaHoy = DateTime.Now.ToString("yyyy-MM-dd");
            string nombreQuitadas = Path.Combine(carpetaQuitadas, $"op quitadas - {fechaHoy}.xlsx");
            string nombreAgregadas = Path.Combine(carpetaAgregadas, $"op agregadas - {fechaHoy}.xlsx");

            var quitadas = new List<object[]>();
            var agregadas = new List<object[]>();

            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                QuitarOperacionesDeSas(rutaSas, fechaCorte, excelApp, quitadas);
                if (quitadas.Count > 0)
                    GuardarExcel(nombreQuitadas, LeerCabecera(rutaSas, excelApp), quitadas, excelApp);

                AgregarDesdeCarpetaQuitadasAlSas(nombreQuitadas, rutaSas, fechaCorte, excelApp, agregadas);
                if (agregadas.Count > 0)
                    GuardarExcel(nombreAgregadas, LeerCabecera(rutaSas, excelApp), agregadas, excelApp);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private void QuitarOperacionesDeSas(string rutaSas, DateTime fechaCorte, Excel.Application excelApp, List<object[]> quitadas)
        {
            var wb = excelApp.Workbooks.Open(rutaSas);
            var hoja = wb.Sheets["Hoja1"] as Excel.Worksheet;
            int lastRow = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            for (int i = lastRow; i >= 2; i--)
            {
                string estado = Convert.ToString((hoja.Cells[i, 16] as Excel.Range)?.Value2)?.Trim();
                string fechaStr = Convert.ToString((hoja.Cells[i, 1] as Excel.Range)?.Value2)?.Trim();

                if (estado == "PENDIENTE-EXEP ANTICIPO" && TryParseFecha(fechaStr, out DateTime fechaA))
                {
                    if (fechaA > fechaCorte)
                    {
                        object[] fila = LeerFila(hoja, i);
                        quitadas.Add(fila);
                        hoja.Rows[i].Delete();
                    }
                }
            }

            wb.Save();
            wb.Close(false);
            Marshal.ReleaseComObject(hoja);
            Marshal.ReleaseComObject(wb);
        }

        private void AgregarDesdeCarpetaQuitadasAlSas(string archivoQuitadas, string rutaSas, DateTime fechaCorte, Excel.Application excelApp, List<object[]> agregadas)
        {
            var carpetaQuitadas = Path.GetDirectoryName(archivoQuitadas);
            var archivos = Directory.GetFiles(carpetaQuitadas, "*.xls*")
                .Where(f => !f.Equals(archivoQuitadas, StringComparison.OrdinalIgnoreCase))
                .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                .Where(f => Path.GetExtension(f).ToLower() is ".xls" or ".xlsx" or ".xlsm")
                .ToList();

            var wbSas = excelApp.Workbooks.Open(rutaSas);
            var hojaSas = wbSas.Sheets["Hoja1"] as Excel.Worksheet;
            string fechaModeloC = Convert.ToString((hojaSas.Cells[2, 3] as Excel.Range)?.Value2);

            foreach (var archivo in archivos)
            {
                try
                {
                    var wb = excelApp.Workbooks.Open(archivo);
                    var hoja = wb.Sheets["Hoja1"] as Excel.Worksheet;
                    int lastRow = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                    bool archivoModificado = false;

                    for (int i = lastRow; i >= 2; i--)
                    {
                        string estado = Convert.ToString((hoja.Cells[i, 16] as Excel.Range)?.Value2)?.Trim();
                        string fechaStr = Convert.ToString((hoja.Cells[i, 1] as Excel.Range)?.Value2)?.Trim();

                        if (estado == "PENDIENTE-EXEP ANTICIPO" && TryParseFecha(fechaStr, out DateTime fechaA))
                        {
                            if (fechaA <= fechaCorte)
                            {
                                object[] fila = LeerFila(hoja, i);
                                fila[2] = fechaModeloC;
                                agregadas.Add(fila);
                                hoja.Rows[i].Delete();
                                archivoModificado = true;
                            }
                        }
                    }

                    if (archivoModificado)
                    {
                        bool tieneSoloCabecera = hoja.UsedRange.Rows.Count <= 1;
                        if (tieneSoloCabecera)
                        {
                            wb.Close(false);
                            Marshal.ReleaseComObject(hoja);
                            Marshal.ReleaseComObject(wb);
                            File.Delete(archivo);
                        }
                        else
                        {
                            wb.Save();
                            wb.Close(false);
                            Marshal.ReleaseComObject(hoja);
                            Marshal.ReleaseComObject(wb);
                        }
                    }
                    else
                    {
                        wb.Close(false);
                        Marshal.ReleaseComObject(hoja);
                        Marshal.ReleaseComObject(wb);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error procesando archivo {archivo}: {ex.Message}");
                }
            }

            int filaInicio = hojaSas.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
            foreach (var fila in agregadas)
            {
                for (int j = 0; j < 22; j++)
                    hojaSas.Cells[filaInicio, j + 1].Value2 = fila[j];
                filaInicio++;
            }

            wbSas.Save();
            wbSas.Close(false);
            Marshal.ReleaseComObject(hojaSas);
            Marshal.ReleaseComObject(wbSas);
        }

        private object[] LeerCabecera(string rutaSas, Excel.Application excelApp)
        {
            var wb = excelApp.Workbooks.Open(rutaSas);
            var hoja = wb.Sheets["Hoja1"] as Excel.Worksheet;
            Excel.Range cabecera = hoja.Range["A1:V1"];
            object[,] valoresCabecera = cabecera.Value2;
            object[] header = Enumerable.Range(1, 22).Select(i => valoresCabecera[1, i]).ToArray();
            wb.Close(false);
            Marshal.ReleaseComObject(hoja);
            Marshal.ReleaseComObject(wb);
            return header;
        }

        private bool TryParseFecha(string raw, out DateTime fecha)
        {
            fecha = default;
            if (DateTime.TryParseExact(raw, "d/M/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                return true;

            if (double.TryParse(raw, out double oa))
            {
                try { fecha = DateTime.FromOADate(oa); return true; } catch { return false; }
            }

            return false;
        }

        private object[] LeerFila(Excel.Worksheet hoja, int fila)
        {
            var rango = hoja.Range[$"A{fila}:V{fila}"].Value2 as object[,];
            return Enumerable.Range(1, 22).Select(i => rango[1, i]).ToArray();
        }

        private void GuardarExcel(string ruta, object[] cabecera, List<object[]> filas, Excel.Application app)
        {
            var wb = app.Workbooks.Add();
            var hoja = wb.Sheets[1] as Excel.Worksheet;

            for (int j = 0; j < 22; j++)
                hoja.Cells[1, j + 1].Value2 = cabecera[j];

            for (int i = 0; i < filas.Count; i++)
            {
                for (int j = 0; j < 22; j++)
                    hoja.Cells[i + 2, j + 1].Value2 = filas[i][j];
            }

            hoja.Columns[1].NumberFormat = "dd/mm/yyyy";
            hoja.Columns[3].NumberFormat = "dd/mm/yyyy";

            wb.SaveAs(ruta);
            wb.Close(false);
            Marshal.ReleaseComObject(hoja);
            Marshal.ReleaseComObject(wb);
        }
    }
}
