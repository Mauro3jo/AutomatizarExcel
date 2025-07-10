using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso2
{
    public class OperacionesDesdeExcepcionService
    {
        public List<List<object>> GenerarFilasDesdeExcepcion(string rutaOriginal, string fechaSeleccionada, Action<string, int>? reportarProgreso = null)
        {
            var filas = new List<List<object>>();
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                reportarProgreso?.Invoke("📂 Abriendo archivo...", 5);

                var workbook = excelApp.Workbooks.Open(rutaOriginal);
                var hojaSAS = workbook.Sheets["SAS"] as Excel.Worksheet;
                var hojaEx = workbook.Sheets["excepcion anticipo"] as Excel.Worksheet;

                string fechaNueva = Convert.ToString((hojaSAS.Cells[2, 3] as Excel.Range)?.Value2);
                int lastRowSAS = hojaSAS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                int ultimoIncremento = 0;
                for (int i = lastRowSAS; i >= 2; i--)
                {
                    var val = Convert.ToString((hojaSAS.Cells[i, 18] as Excel.Range)?.Value2);
                    if (int.TryParse(val ?? "", out ultimoIncremento))
                        break;
 
                
                }

                int lastRowEx = hojaEx.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                int contador = 0;
                for (int i = 2; i <= lastRowEx; i++)
                {
                    reportarProgreso?.Invoke($"🔎 Buscando filas coincidentes... ({i}/{lastRowEx})", i * 100 / lastRowEx);

                    var celdaZ = hojaEx.Cells[i, 26] as Excel.Range;
                    string fechaZ = "";

                    if (celdaZ?.Value2 != null)
                    {
                        try
                        {
                            fechaZ = DateTime.FromOADate(Convert.ToDouble(celdaZ.Value2)).ToString("d/M/yyyy");
                        }
                        catch
                        {
                            fechaZ = celdaZ.Value2.ToString().Trim();
                        }
                    }

                    if (fechaZ != fechaSeleccionada)
                        continue;

                    Excel.Range filaRango = hojaEx.Range[$"A{i}:V{i}"];
                    object[,] valores = filaRango.Value2 as object[,];

                    var nuevaFila = new List<object>();
                    for (int col = 0; col < 22; col++)
                    {
                        object valor = valores[1, col + 1] ?? "";

                        if (col + 1 == 9) valor = ""; // I
                        if (col + 1 == 3) valor = fechaNueva; // C
                        if (col + 1 == 16) valor = "PENDIENTE-EXEP ANTICIPO"; // P
                        if (col + 1 == 18) valor = (++ultimoIncremento).ToString(); // R

                        nuevaFila.Add(valor);
                    }

                    filas.Add(nuevaFila);
                    contador++;
                }

                workbook.Close(false);
                Marshal.ReleaseComObject(hojaSAS);
                Marshal.ReleaseComObject(hojaEx);
                Marshal.ReleaseComObject(workbook);

                reportarProgreso?.Invoke($"✅ Se encontraron {contador} filas para insertar.", 100);
            }
            catch (Exception ex)
            {
                reportarProgreso?.Invoke("❌ Error leyendo datos desde excepción: " + ex.Message, 0);
                throw;
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }

            return filas;
        }

        public void AgregarFilasAlSAS(string rutaArchivo, List<List<object>> filas, Action<string, int>? reportarProgreso = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            try
            {
                reportarProgreso?.Invoke("📄 Abriendo archivo destino...", 5);

                var workbook = excelApp.Workbooks.Open(rutaArchivo);
                var hojaSAS = workbook.Sheets["Hoja1"] as Excel.Worksheet;

                int rowDestino = hojaSAS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                int total = filas.Count;

                for (int i = 0; i < total; i++)
                {
                    if (filas[i].Count != 22)
                        throw new Exception($"Fila en índice {i} no tiene 22 columnas, tiene {filas[i].Count}");

                    Excel.Range filaDestino = hojaSAS.Range[$"A{rowDestino + i}:V{rowDestino + i}"];
                    object[,] filaData = new object[1, 22];

                    for (int j = 0; j < 22; j++)
                        filaData[0, j] = filas[i][j];

                    filaDestino.Value2 = filaData;

                    int progreso = (int)((i + 1) / (float)total * 100);
                    reportarProgreso?.Invoke($"✍️ Insertando fila {i + 1} de {total}...", progreso);
                }

                workbook.Save();
                workbook.Close(false);

                Marshal.ReleaseComObject(hojaSAS);
                Marshal.ReleaseComObject(workbook);

                reportarProgreso?.Invoke("✅ Agregado completado correctamente.", 100);
            }
            catch (Exception ex)
            {
                reportarProgreso?.Invoke("❌ Error agregando filas: " + ex.Message, 0);
                throw;
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
