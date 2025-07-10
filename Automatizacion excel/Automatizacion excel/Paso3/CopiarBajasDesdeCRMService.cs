using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso3
{
    public class CopiarBajasDesdeCRMService
    {
        public void CopiarBajas(string rutaCRM, string rutaCrudo, Action<string, int>? reportarProgreso = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;

            Excel.Workbook wbCRM = null;
            Excel.Workbook wbCrudo = null;

            try
            {
                reportarProgreso?.Invoke("📂 Abriendo archivos...", 5);

                wbCRM = excelApp.Workbooks.Open(rutaCRM);
                wbCrudo = excelApp.Workbooks.Open(rutaCrudo);

                var hojaCRM = wbCRM.Sheets["BAJAS"] as Excel.Worksheet;
                var hojaCrudo = wbCrudo.Sheets["Bajas"] as Excel.Worksheet;

                if (hojaCRM == null || hojaCrudo == null)
                    throw new Exception("No se encontraron las hojas BAJAS o Bajas.");

                int lastRowCRM = hojaCRM.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                int colHasta = 68; // A = 1, BP = 68
                int filasCopiadas = 0;

                for (int fila = 2; fila <= lastRowCRM; fila++)
                {
                    Excel.Range origen = hojaCRM.Range[hojaCRM.Cells[fila, 1], hojaCRM.Cells[fila, colHasta]];
                    Excel.Range destino = hojaCrudo.Range[hojaCrudo.Cells[fila, 1], hojaCrudo.Cells[fila, colHasta]];
                    destino.Value2 = origen.Value2;

                    filasCopiadas++;
                    int progreso = (int)((fila - 1) / (float)(lastRowCRM - 1) * 100);
                    reportarProgreso?.Invoke($"✍️ Copiando fila {fila} de {lastRowCRM}...", progreso);
                }

                wbCrudo.Save();
                reportarProgreso?.Invoke($"✅ {filasCopiadas} bajas copiadas correctamente.", 100);
            }
            catch (Exception ex)
            {
                reportarProgreso?.Invoke("❌ Error al copiar bajas: " + ex.Message, 0);
                throw;
            }
            finally
            {
                wbCRM?.Close(false);
                wbCrudo?.Close(false);
                excelApp.Quit();

                Marshal.ReleaseComObject(wbCRM);
                Marshal.ReleaseComObject(wbCrudo);
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
