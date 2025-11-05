using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Runtime.InteropServices;

namespace Automatizacion_excel.Paso2
{
    public class ProcesarExcepcionAnticipoService
    {
        private readonly string connectionString;

        public ProcesarExcepcionAnticipoService()
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            connectionString = config.GetConnectionString("MiConexion");
        }

        public void EjecutarProceso(string rutaExcel, Action<string, int>? reportar = null)
        {
            var excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbook wb = null;
            Excel.Worksheet hojaPrincipal = null;
            Excel.Worksheet hojaQuitadas = null;

            try
            {
                wb = excelApp.Workbooks.Open(rutaExcel);
                hojaPrincipal = wb.Sheets["Hoja1"];
                hojaQuitadas = wb.Sheets["Op Quitadas"];

                int lastRowHoja1 = hojaPrincipal.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                reportar?.Invoke("🔍 Analizando terminales activas...", 10);

                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();

                    // --- 1) De Excel → Base y “Op Quitadas”
                    for (int i = lastRowHoja1; i >= 2; i--)
                    {
                        string nroComercio = Convert.ToString((hojaPrincipal.Cells[i, 5] as Excel.Range)?.Value2)?.Trim();
                        if (string.IsNullOrEmpty(nroComercio)) continue;

                        if (TerminalActiva(conn, nroComercio))
                        {
                            var fila = LeerFila(hojaPrincipal, i);
                            int.TryParse(fila[16]?.ToString(), out int cuotas);
                            string tarjeta = fila[18]?.ToString()?.Trim().ToUpper() ?? "";

                            // ❌ No se procesan ni se quitan las que no sean VISA, MASTER, ARGENCARD o sean cuota 0
                            if (cuotas == 0 ||
                                !(tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD")))
                                continue;

                            // Copiar al final de “Op Quitadas”
                            int lastRowQuitadas = hojaQuitadas.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                            PegarFila(hojaQuitadas, lastRowQuitadas, fila);

                            // Insertar en base solo si cumple condiciones válidas
                            InsertarEnBase(conn, fila);

                            // Eliminar de Hoja1
                            hojaPrincipal.Rows[i].Delete();
                        }
                    }

                    // --- 2) De Base → Excel
                    reportar?.Invoke("📥 Agregando operaciones desde la base a Hoja1...", 60);
                    AgregarPendientesAHoja1(conn, hojaPrincipal);

                    conn.Close();
                }

                wb.Save();
                reportar?.Invoke("✅ Proceso completado correctamente y guardado.", 100);
            }
            finally
            {
                if (hojaPrincipal != null) Marshal.ReleaseComObject(hojaPrincipal);
                if (hojaQuitadas != null) Marshal.ReleaseComObject(hojaQuitadas);
                if (wb != null)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }

        // ---------- MÉTODOS PRINCIPALES ----------

        private bool TerminalActiva(SqlConnection conn, string nroTerminal)
        {
            const string sql = @"SELECT COUNT(*) FROM [dbo].[TerminalesExcepcionAnticipo]
                                 WHERE NroTerminal = @NroTerminal AND EstaEliminado = 0";
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@NroTerminal", nroTerminal);
                return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
            }
        }

        private void InsertarEnBase(SqlConnection conn, object[] fila)
        {
            DateTime fechaOperacion = ParsearFecha(fila[0]) ?? DateTime.MinValue;
            DateTime fechaPago = ParsearFecha(fila[2]) ?? DateTime.MinValue;
            string tarjeta = fila[18]?.ToString()?.Trim().ToUpper() ?? "";

            int.TryParse(fila[16]?.ToString(), out int cuotas);

            // ❌ Solo se insertan VISA, MASTER y ARGENCARD con cuotas > 0
            if (cuotas == 0 ||
                !(tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD")))
                return;

            string tipoPago = cuotas == 1 ? "CRÉDITO 1 PAGO" : "CRÉDITO 2 O MÁS PAGOS";
            string categoriaTarjeta = DetectarCategoriaTarjeta(tarjeta);

            int dias = ObtenerDiasPlazo(conn, tipoPago, categoriaTarjeta, fila[8]);
            DateTime fechaAAgregar = SumarDiasHabiles(fechaOperacion, dias);

            const string sql = @"
                INSERT INTO [dbo].[ExcepcionAnticipo]
                ([FechaOperacion],[FechaPresentacion],[FechaPago],[NroCupon],
                 [NroComercio],[NroTarjeta],[Moneda],[TotalBruto],[TotalDescuento],[TotalNeto],
                 [EntidadPagadora],[CuentaBancaria],[NroLiquidacion],[NroLote],[TipoLiquidacion],
                 [Estado],[Cuotas],[NroAutorizacion],[Tarjeta],[TipoOperacion],[ComercioParticipante],
                 [PromocionPlan],[DiasAdelanto],[FechaAAgregar],[EstadoNuevo],[PagoOriginal],[FechaPagado])
                VALUES
                (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13,@p14,
                 @p15,@p16,@p17,@p18,@p19,@p20,@p21,@DiasAdelanto,@FechaAAgregar,@EstadoNuevo,@PagoOriginal,@FechaPagado);";

            using (var cmd = new SqlCommand(sql, conn))
            {
                for (int i = 0; i < 22; i++)
                {
                    object valor = fila[i] ?? DBNull.Value;
                    if (i == 0 || i == 1 || i == 2)
                        valor = ParsearFecha(valor) ?? (object)DBNull.Value;
                    else if (i == 7 || i == 8 || i == 9)
                        valor = LimpiarDecimal(valor);
                    cmd.Parameters.AddWithValue($"@p{i}", valor);
                }

                cmd.Parameters.AddWithValue("@DiasAdelanto", dias);
                cmd.Parameters.AddWithValue("@FechaAAgregar", fechaAAgregar);
                cmd.Parameters.AddWithValue("@EstadoNuevo", "NO PAGADO");
                cmd.Parameters.AddWithValue("@PagoOriginal", fechaPago);
                cmd.Parameters.AddWithValue("@FechaPagado", DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        private void AgregarPendientesAHoja1(SqlConnection conn, Excel.Worksheet hoja)
        {
            object valorC2 = (hoja.Cells[2, 3] as Excel.Range)?.Value2;
            DateTime? fechaPagoModelo = ParsearFecha(valorC2);
            string fechaPagoTexto = fechaPagoModelo?.ToString("dd/MM/yyyy");

            const string sql = @"
        SELECT FechaOperacion, FechaPresentacion, FechaPago, NroCupon, NroComercio, NroTarjeta,
               Moneda, TotalBruto, TotalDescuento, TotalNeto, EntidadPagadora, CuentaBancaria,
               NroLiquidacion, NroLote, TipoLiquidacion, Estado, Cuotas, NroAutorizacion,
               Tarjeta, TipoOperacion, ComercioParticipante, PromocionPlan
        FROM [dbo].[ExcepcionAnticipo]
        WHERE EstadoNuevo = 'NO PAGADO' 
        AND FechaAAgregar <= GETDATE()
        AND ISNULL(Cuotas, 0) > 0
        ORDER BY ID";

            using (var cmd = new SqlCommand(sql, conn))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    int nextRow = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                    for (int col = 1; col <= 22; col++)
                    {
                        object valor = reader.GetValue(col - 1);
                        if (col == 3)
                            valor = fechaPagoTexto;
                        else if (col == 9)
                            valor = "";
                        else if (col == 16)
                            valor = "PENDIENTE-EXEP ANTICIPO";
                        hoja.Cells[nextRow, col].Value2 = valor;
                    }
                }
            }

            const string updateSql = @"
        UPDATE [dbo].[ExcepcionAnticipo]
        SET EstadoNuevo = 'PAGADO', FechaPagado = GETDATE()
        WHERE EstadoNuevo = 'NO PAGADO'
        AND FechaAAgregar <= GETDATE()
        AND ISNULL(Cuotas, 0) > 0;";
            using (var cmd = new SqlCommand(updateSql, conn))
            {
                cmd.ExecuteNonQuery();
            }
        }

        // ---------- UTILIDADES ----------

        private void PegarFila(Excel.Worksheet hojaDestino, int filaDestino, object[] filaOrigen)
        {
            for (int col = 0; col < filaOrigen.Length; col++)
                hojaDestino.Cells[filaDestino, col + 1].Value2 = filaOrigen[col];
        }

        private object[] LeerFila(Excel.Worksheet hoja, int fila)
        {
            var rango = hoja.Range[$"A{fila}:V{fila}"].Value2 as object[,];
            object[] arr = new object[22];
            for (int i = 0; i < 22; i++)
                arr[i] = rango[1, i + 1];
            return arr;
        }

        private DateTime? ParsearFecha(object valor)
        {
            if (valor == null || valor == DBNull.Value) return null;
            try
            {
                if (valor is DateTime dt) return dt;
                string s = valor.ToString().Trim();
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
                    return DateTime.FromOADate(oa);
                if (DateTime.TryParse(s, new CultureInfo("es-AR"), DateTimeStyles.None, out dt))
                    return dt;
                return null;
            }
            catch { return null; }
        }

        private string DetectarCategoriaTarjeta(string tarjeta)
        {
            tarjeta = tarjeta.ToUpperInvariant();
            if (tarjeta.Contains("CABAL")) return "Bancarizadas (Cabal)";
            if (tarjeta.Contains("AMEX")) return "Bancarizadas (Amex)";
            if (tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD"))
                return "Bancarizadas (Visa - Master - ArgenCard)";
            if (tarjeta.Contains("NARANJA"))
                return "No Bancarizadas (Naranja Visa, Naranja Master, Cencosud, etc.)";
            return "";
        }

        private int ObtenerDiasPlazo(SqlConnection conn, string tipoPago, string categoriaTarjeta, object totalDescuentoObj)
        {
            decimal totalDescuento = 0;
            if (totalDescuentoObj != null && decimal.TryParse(
                    totalDescuentoObj.ToString().Replace("%", "").Replace(",", ".").Trim(),
                    NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                totalDescuento = val;

            // 🟡 Excepción: Crédito 1 pago con >4% descuento
            if (tipoPago == "CRÉDITO 1 PAGO" && totalDescuento > 4)
                return 17;

            // 🔵 Caso normal
            const string sql = @"SELECT TOP 1 Dias FROM [dbo].[PlazosDeAcreditaciones]
                                 WHERE TipoPago = @TipoPago AND Tarjetas LIKE '%' + @Tarjetas + '%'";
            using (var cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.AddWithValue("@TipoPago", tipoPago);
                cmd.Parameters.AddWithValue("@Tarjetas", categoriaTarjeta);
                var result = cmd.ExecuteScalar();
                return result != null ? Convert.ToInt32(result) : 0;
            }
        }

        private DateTime SumarDiasHabiles(DateTime fechaInicio, int dias)
        {
            int agregados = 0;
            DateTime fecha = fechaInicio.AddDays(1); // empieza al día siguiente
            while (agregados < dias)
            {
                if (fecha.DayOfWeek != DayOfWeek.Saturday && fecha.DayOfWeek != DayOfWeek.Sunday)
                    agregados++;
                if (agregados < dias)
                    fecha = fecha.AddDays(1);
            }
            return fecha;
        }

        private object LimpiarDecimal(object valor)
        {
            if (valor == null || valor == DBNull.Value) return DBNull.Value;
            string str = valor.ToString()
                .Replace("$", "")
                .Replace("%", "")
                .Replace(" ", "")
                .Replace(".", "")
                .Replace(",", ".")
                .Trim();
            if (decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal num))
                return num;
            return DBNull.Value;
        }
    }
}
