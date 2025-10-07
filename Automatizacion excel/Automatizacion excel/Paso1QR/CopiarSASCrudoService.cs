using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso1QR
{
    public class CopiarSASCrudoService
    {
        public void CopiarSAS(string rutaConversor, string rutaCrudo, Action<string, int>? reportarProgreso = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            Excel.Workbook wbConversor = null;
            Excel.Workbook wbCrudo = null;

            try
            {
                reportarProgreso?.Invoke("📂 Abriendo archivos...", 5);

                wbConversor = excelApp.Workbooks.Open(rutaConversor);
                wbCrudo = excelApp.Workbooks.Open(rutaCrudo);

                var hojaSAS = wbConversor.Sheets["SAS"] as Excel.Worksheet;
                var hojaDestino = wbCrudo.Sheets["crudo"] as Excel.Worksheet;

                if (hojaSAS == null || hojaDestino == null)
                    throw new Exception("No se encontró la hoja 'SAS' en el Conversor o 'crudo' en el archivo destino.");

                int lastRow = hojaSAS.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int colFin = 22; // Columna V

                // Limpiar contenido anterior
                hojaDestino.Range["A2", hojaDestino.Cells[hojaDestino.Rows.Count, colFin]].ClearContents();

                // Copiar datos sin tocar fechas
                var rangoOrigen = hojaSAS.Range[hojaSAS.Cells[2, 1], hojaSAS.Cells[lastRow, colFin]];
                var rangoDestino = hojaDestino.Range[hojaDestino.Cells[2, 1], hojaDestino.Cells[lastRow, colFin]];
                rangoDestino.Value = rangoOrigen.Value;

                reportarProgreso?.Invoke("📄 Datos copiados. Iniciando validación de fechas...", 40);

                // ===============================
                // 🔧 Revisar y corregir fechas directamente en 'crudo'
                // ===============================
                int filaActual = 2;
                while (true)
                {
                    var celdaOperacion = hojaDestino.Cells[filaActual, 1];
                    var celdaPago = hojaDestino.Cells[filaActual, 3];

                    if (celdaOperacion == null || celdaPago == null)
                        break;

                    var valOperacion = celdaOperacion.Value;
                    var valPago = celdaPago.Value;

                    if (valOperacion == null && valPago == null)
                        break; // fin de datos

                    if (valOperacion == null || valPago == null)
                    {
                        filaActual++;
                        continue;
                    }

                    try
                    {
                        DateTime fechaPago = Convert.ToDateTime(valPago);
                        DateTime fechaOperacion = Convert.ToDateTime(valOperacion);

                        bool mismoMes = fechaOperacion.Month == fechaPago.Month && fechaOperacion.Year == fechaPago.Year;
                        bool mesAnterior = fechaOperacion.Year == fechaPago.Year &&
                                           fechaOperacion.Month == (fechaPago.Month == 1 ? 12 : fechaPago.Month - 1);
                        bool diaValido = fechaOperacion.Day < fechaPago.Day;

                        // ✅ Si cumple, dejar igual
                        if ((mismoMes || mesAnterior) && diaValido)
                        {
                            filaActual++;
                            continue;
                        }

                        // 🔁 Intentar invertir
                        bool invertidaValida = false;
                        string texto = valOperacion.ToString();

                        try
                        {
                            string[] partes = texto.Split('/');
                            if (partes.Length == 3 &&
                                int.TryParse(partes[0], out int p1) &&
                                int.TryParse(partes[1], out int p2) &&
                                int.TryParse(partes[2], out int p3))
                            {
                                DateTime invertida = new DateTime(p3, p1, p2);

                                bool mm = invertida.Month == fechaPago.Month;
                                bool ma = invertida.Month == (fechaPago.Month == 1 ? 12 : fechaPago.Month - 1);
                                bool dv = invertida.Day < fechaPago.Day;

                                if ((mm || ma) && dv)
                                {
                                    celdaOperacion.Value = invertida;
                                    invertidaValida = true;
                                }
                            }
                        }
                        catch { }

                        if (invertidaValida)
                        {
                            filaActual++;
                            continue;
                        }

                        // ⚙️ Ajustar al día anterior del pago
                        int nuevoDia = fechaPago.Day - 1;
                        if (nuevoDia <= 0)
                            nuevoDia = 1;

                        DateTime fechaAjustada = new DateTime(fechaPago.Year, fechaPago.Month, nuevoDia);
                        celdaOperacion.Value = fechaAjustada;
                    }
                    catch { }

                    filaActual++;
                }

                // 🧾 Aplicar formato visual a la columna de fecha
                var rangoFechas = hojaDestino.Range["A2", hojaDestino.Cells[lastRow, 1]];
                rangoFechas.NumberFormat = "dd/mm/yyyy";  // <--- 🔥 clave

                wbCrudo.Save();
                reportarProgreso?.Invoke("✅ Fechas revisadas y formato aplicado correctamente en el Crudo.", 100);
            }
            catch (Exception ex)
            {
                reportarProgreso?.Invoke("❌ Error al copiar desde SAS: " + ex.Message, 0);
                throw;
            }
            finally
            {
                wbConversor?.Close(false);
                wbCrudo?.Close(true);
                excelApp.Quit();

                Marshal.ReleaseComObject(wbConversor);
                Marshal.ReleaseComObject(wbCrudo);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
