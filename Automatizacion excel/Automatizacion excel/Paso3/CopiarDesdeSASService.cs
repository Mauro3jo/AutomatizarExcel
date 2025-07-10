using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso3
{
    public class CopiarDesdeSASService
    {
        public void CopiarSAS(string rutaSAS, string rutaCrudo, Action<string, int>? reportarProgreso = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            Excel.Workbook wbSAS = null;
            Excel.Workbook wbCrudo = null;

            try
            {
                reportarProgreso?.Invoke("📂 Abriendo archivos...", 5);

                wbSAS = excelApp.Workbooks.Open(rutaSAS);
                wbCrudo = excelApp.Workbooks.Open(rutaCrudo);

                var hojaSasOrigen = wbSAS.Sheets["Hoja1"] as Excel.Worksheet;
                var hojaDestino = wbCrudo.Sheets["crudo"] as Excel.Worksheet;

                if (hojaSasOrigen == null || hojaDestino == null)
                    throw new Exception("No se encontró la hoja 'Hoja1' en el SAS o 'crudo' en el archivo destino.");

                int lastRow = hojaSasOrigen.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int colFin = 22;

                // ✅ Borrar contenido anterior (desde fila 2)
                hojaDestino.Range["A2", hojaDestino.Cells[hojaDestino.Rows.Count, colFin]].ClearContents();

                // 🧠 Definir rangos de origen y destino
                var rangoOrigen = hojaSasOrigen.Range[
                    hojaSasOrigen.Cells[2, 1],
                    hojaSasOrigen.Cells[lastRow, colFin]
                ];

                var valores = rangoOrigen.Value;

                var rangoDestino = hojaDestino.Range[
                    hojaDestino.Cells[2, 1],
                    hojaDestino.Cells[lastRow, colFin]
                ];

                // 🏎️ Pegar todo en una sola operación
                rangoDestino.Value = valores;

                wbCrudo.Save();
                reportarProgreso?.Invoke("✅ SAS copiado correctamente al Crudo.", 100);
            }
            catch (Exception ex)
            {
                reportarProgreso?.Invoke("❌ Error al copiar desde SAS: " + ex.Message, 0);
                throw;
            }
            finally
            {
                wbSAS?.Close(false);
                wbCrudo?.Close(false);
                excelApp.Quit();

                Marshal.ReleaseComObject(wbSAS);
                Marshal.ReleaseComObject(wbCrudo);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
