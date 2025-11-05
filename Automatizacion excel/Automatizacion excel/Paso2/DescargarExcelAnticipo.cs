using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Configuration;

namespace Automatizacion_excel.Paso2
{
    public class DescargarExcelAnticipo
    {
        private readonly string connectionString;

        public DescargarExcelAnticipo()
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            connectionString = config.GetConnectionString("MiConexion");
        }

        public void Ejecutar()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel|*.xlsx";
            sfd.Title = "Guardar Excel de anticipos";
            sfd.FileName = "ExcepcionAnticipo_" + DateTime.Now.ToString("yyyyMMdd") + ".xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                DataTable dt = ObtenerDatos();

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("No se encontraron registros en la tabla [ExcepcionAnticipo].", "Sin datos");
                    return;
                }

                ExportarAExcel(dt, sfd.FileName);
                MessageBox.Show("Archivo exportado correctamente a:\n" + sfd.FileName, "¡Éxito!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar Excel:\n" + ex.Message, "Error");
            }
        }

        private DataTable ObtenerDatos()
        {
            using (var conn = new SqlConnection(connectionString))
            using (var cmd = new SqlCommand("SELECT * FROM [dbo].[ExcepcionAnticipo] ORDER BY ID", conn))
            using (var da = new SqlDataAdapter(cmd))
            {
                var dt = new DataTable();
                conn.Open();
                da.Fill(dt);
                return dt;
            }
        }

        private void ExportarAExcel(DataTable dt, string ruta)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            var wb = excelApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet ws = wb.ActiveSheet;
            ws.Name = "ExcepcionAnticipo";

            // Escribir encabezados
            for (int i = 0; i < dt.Columns.Count; i++)
                ws.Cells[1, i + 1] = dt.Columns[i].ColumnName;

            // Escribir filas
            for (int r = 0; r < dt.Rows.Count; r++)
                for (int c = 0; c < dt.Columns.Count; c++)
                    ws.Cells[r + 2, c + 1] = dt.Rows[r][c];

            ws.Columns.AutoFit();
            wb.SaveAs(ruta);
            wb.Close();
            excelApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}
