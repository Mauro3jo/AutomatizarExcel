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
            // 📂 Carga la cadena de conexión desde appsettings.json
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            connectionString = config.GetConnectionString("MiConexion");
        }

        // ---------------------------------------------------------------------
        // 🔹 Método principal que ejecuta todo el flujo
        // ---------------------------------------------------------------------
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

                    // --- 🔹 1) Recorre el Excel y analiza las filas válidas
                    for (int i = lastRowHoja1; i >= 2; i--)
                    {
                        string nroComercio = Convert.ToString((hojaPrincipal.Cells[i, 5] as Excel.Range)?.Value2)?.Trim();
                        if (string.IsNullOrEmpty(nroComercio)) continue;

                        // Verifica si la terminal está activa
                        if (TerminalActiva(conn, nroComercio))
                        {
                            var fila = LeerFila(hojaPrincipal, i);
                            int.TryParse(fila[16]?.ToString(), out int cuotas);
                            string tarjeta = fila[18]?.ToString()?.Trim().ToUpper() ?? "";

                            // ❌ No procesar ni eliminar:
                            // - Cabal, Amex, Maestro, Naranja o Cuota 0 (débito)
                            // ✅ Solo Visa, Master o ArgenCard con cuotas > 0
                            if (cuotas == 0 ||
                                !(tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD")))
                                continue;

                            // 📋 Copiar al final de la hoja “Op Quitadas”
                            int lastRowQuitadas = hojaQuitadas.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                            PegarFila(hojaQuitadas, lastRowQuitadas, fila);

                            // 💾 Insertar en base con la lógica de días y fechas
                            InsertarEnBase(conn, fila);

                            // ❌ Eliminar de la hoja principal
                            hojaPrincipal.Rows[i].Delete();
                        }
                    }

                    // --- 🔹 2) Cargar operaciones pendientes desde la base al Excel
                    reportar?.Invoke("📥 Agregando operaciones desde la base a Hoja1...", 60);
                    AgregarPendientesAHoja1(conn, hojaPrincipal);

                    conn.Close();
                }

                // ✅ Guardar los cambios
                wb.Save();
                reportar?.Invoke("✅ Proceso completado correctamente y guardado.", 100);
            }
            finally
            {
                // Limpieza de objetos COM
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

        // ---------------------------------------------------------------------
        // 🔹 Verifica si la terminal está activa en la tabla
        // ---------------------------------------------------------------------
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

        // ---------------------------------------------------------------------
        // 🔹 Inserta una operación válida en ExcepcionAnticipo
        // ---------------------------------------------------------------------
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

            // 🧮 Calcular descuento correctamente (detecta "8,66%" o "0,0866")
            decimal totalDescuento = 0;
            if (fila[8] != null)
            {
                string s = fila[8].ToString().Replace("%", "").Trim().Replace(",", ".");
                if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                {
                    if (val > 1) val /= 100; // si viene como 8.66 => 0.0866
                    totalDescuento = val;
                }
            }

            int dias;
            if (cuotas >= 2)
                dias = 9; // 2 o más cuotas
            else if (cuotas == 1 && totalDescuento > 0.04m)
                dias = 17; // 1 pago con descuento > 4%
            else
                dias = 7; // 1 pago con descuento ≤ 4%

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

        // ---------------------------------------------------------------------
        // 🔹 Agrega operaciones pendientes desde la base a la hoja principal
        // ---------------------------------------------------------------------
        private void AgregarPendientesAHoja1(SqlConnection conn, Excel.Worksheet hoja)
        {
            object valorC2 = (hoja.Cells[2, 3] as Excel.Range)?.Value2;
            DateTime? fechaPagoModelo = ParsearFecha(valorC2);
            string fechaPagoTexto = fechaPagoModelo?.ToString("dd/MM/yyyy") ?? "";

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
                    // 🧩 Busca la primera fila vacía (sin fecha operación ni fecha pago)
                    int nextRow = 2;
                    while (true)
                    {
                        var celdaOp = (hoja.Cells[nextRow, 1] as Excel.Range)?.Value2;
                        var celdaPago = (hoja.Cells[nextRow, 3] as Excel.Range)?.Value2;
                        if (celdaOp == null && celdaPago == null)
                            break;
                        nextRow++;
                    }

                    // 📋 Pega los valores en la fila encontrada
                    for (int col = 1; col <= 22; col++)
                    {
                        object valor = reader.GetValue(col - 1);

                        if (col == 3)
                        {
                            if (DateTime.TryParse(valor?.ToString(), out DateTime fechaPago))
                                valor = fechaPago.ToString("dd/MM/yyyy");
                            else
                                valor = fechaPagoTexto;
                            hoja.Cells[nextRow, col].NumberFormat = "@"; // formato texto
                            hoja.Cells[nextRow, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            hoja.Cells[nextRow, col].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                        }
                        else if (col == 9)
                        {
                            valor = "";
                        }
                        else if (col == 16)
                        {
                            valor = "PENDIENTE-EXCEP ANTICIPO";
                        }

                        hoja.Cells[nextRow, col].Value2 = valor;
                    }
                }
            }

            // 🔄 Actualiza el estado de las operaciones vencidas a PAGADO
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

        // ---------------------------------------------------------------------
        // 🔹 Calcula fecha hábil sumando días (sin fines de semana)
        // ---------------------------------------------------------------------
        private DateTime SumarDiasHabiles(DateTime fechaInicio, int dias)
        {
            int agregados = 0;
            DateTime fecha = fechaInicio.AddDays(1);
            while (agregados < dias)
            {
                if (fecha.DayOfWeek != DayOfWeek.Saturday && fecha.DayOfWeek != DayOfWeek.Sunday)
                    agregados++;
                if (agregados < dias)
                    fecha = fecha.AddDays(1);
            }
            return fecha;
        }

        // ---------------------------------------------------------------------
        // 🔹 Utilidades auxiliares
        // ---------------------------------------------------------------------
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
