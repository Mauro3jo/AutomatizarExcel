using System;
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
                int colFin = 22; // V = 22

                // Borrar contenido anterior (desde fila 2)
                hojaDestino.Range["A2", hojaDestino.Cells[hojaDestino.Rows.Count, colFin]].ClearContents();

                // Definir rangos de origen y destino
                var rangoOrigen = hojaSAS.Range[hojaSAS.Cells[2, 1], hojaSAS.Cells[lastRow, colFin]];
                var valores = rangoOrigen.Value;
                var rangoDestino = hojaDestino.Range[hojaDestino.Cells[2, 1], hojaDestino.Cells[lastRow, colFin]];
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
                wbConversor?.Close(false);
                wbCrudo?.Close(false);
                excelApp.Quit();

                Marshal.ReleaseComObject(wbConversor);
                Marshal.ReleaseComObject(wbCrudo);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
