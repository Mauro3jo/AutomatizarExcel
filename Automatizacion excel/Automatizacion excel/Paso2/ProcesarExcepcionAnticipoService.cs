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

                            // 📋 PEGADO ELIMINADO: antes duplicaba en Op Quitadas
                            //int lastRowQuitadas = hojaQuitadas.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
                            //PegarFila(hojaQuitadas, lastRowQuitadas, fila);

                            // 💾 Insertar en base con la lógica de días y fechas
                            InsertarEnBase(conn, fila, hojaPrincipal, hojaQuitadas, i);

                            // ❌ ELIMINADO: esta línea borraba filas de más
                            //hojaPrincipal.Rows[i].Delete();
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
        // 🔹 Inserta una operación válida en ExcepcionAnticipo o marca duplicado
        // ---------------------------------------------------------------------
        private void InsertarEnBase(SqlConnection conn, object[] fila, Excel.Worksheet hojaPrincipal, Excel.Worksheet hojaQuitadas, int filaIndice)
        {
            // 🔹 Parseo de datos base
            DateTime fechaOperacion = ParsearFecha(fila[0]) ?? DateTime.MinValue;
            DateTime fechaPago = ParsearFecha(fila[2]) ?? DateTime.MinValue;
            string nroCupon = fila[3]?.ToString()?.Trim() ?? "";
            string nroComercio = fila[4]?.ToString()?.Trim() ?? "";
            string nroTarjeta = fila[5]?.ToString()?.Trim() ?? "";
            string tarjeta = fila[18]?.ToString()?.Trim().ToUpper() ?? "";

            int.TryParse(fila[16]?.ToString(), out int cuotas);

            // ❌ No procesar tarjetas no válidas o sin cuotas
            if (cuotas == 0 ||
                !(tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD")))
                return;

            // 🔹 Verificar si ya existe en la base
            const string existeSql = @"
        SELECT COUNT(*) FROM [dbo].[ExcepcionAnticipo]
        WHERE FechaOperacion = @FechaOperacion
          AND NroCupon = @NroCupon
          AND NroComercio = @NroComercio
          AND NroTarjeta = @NroTarjeta;";

            bool yaExiste;
            using (var checkCmd = new SqlCommand(existeSql, conn))
            {
                checkCmd.Parameters.AddWithValue("@FechaOperacion", fechaOperacion);
                checkCmd.Parameters.AddWithValue("@NroCupon", nroCupon);
                checkCmd.Parameters.AddWithValue("@NroComercio", nroComercio);
                checkCmd.Parameters.AddWithValue("@NroTarjeta", nroTarjeta);
                yaExiste = Convert.ToInt32(checkCmd.ExecuteScalar()) > 0;
            }

            // 🔸 Siempre copiar al final de “Op Quitadas” y eliminar de “Hoja1”
            int lastRowQuitadas = hojaQuitadas.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row + 1;
            for (int c = 1; c <= fila.Length; c++)
                hojaQuitadas.Cells[lastRowQuitadas, c].Value2 = fila[c - 1];

            hojaPrincipal.Rows[filaIndice].Delete();

            // 🔸 Si ya existe, no hacer nada más
            if (yaExiste)
                return;

            // 🔹 Calcular descuento y días hábiles
            decimal totalDescuento = 0;
            if (fila[8] != null)
            {
                string s = fila[8].ToString().Replace("%", "").Trim().Replace(",", ".");
                if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                {
                    if (val > 1) val /= 100;
                    totalDescuento = val;
                }
            }

            int dias = (cuotas >= 2) ? 9 :
                       (cuotas == 1 && totalDescuento > 0.045m) ? 17 : 7;

            DateTime fechaAAgregar = SumarDiasHabiles(fechaOperacion, dias);

            // 🔹 Insertar nuevo registro en base
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
        // 🔹 Agrega operaciones pendientes desde la base al Excel
        // ---------------------------------------------------------------------
        private void AgregarPendientesAHoja1(SqlConnection conn, Excel.Worksheet hoja)
        {
            object valorC2 = (hoja.Cells[2, 3] as Excel.Range)?.Value2;
            DateTime? fechaPagoModelo = ParsearFecha(valorC2);

            const string sql = @"
SELECT ID, FechaOperacion, FechaPresentacion, FechaPago, NroCupon, NroComercio, NroTarjeta,
       Moneda, TotalBruto, TotalDescuento, TotalNeto, EntidadPagadora, CuentaBancaria,
       NroLiquidacion, NroLote, TipoLiquidacion, Estado, Cuotas, NroAutorizacion,
       Tarjeta, TipoOperacion, ComercioParticipante, PromocionPlan, FechaAAgregar, EstadoNuevo
FROM [dbo].[ExcepcionAnticipo]
WHERE FechaAAgregar <= CAST(GETDATE() AS date)
  AND ISNULL(Cuotas, 0) > 0
  AND (EstadoNuevo IS NULL OR EstadoNuevo = 'NO PAGADO')
ORDER BY FechaOperacion, NroCupon, NroComercio, NroTarjeta, ID;";

            DataTable dt = new DataTable();
            using (var adapter = new SqlDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
            }

            var regs = dt.AsEnumerable()
                .Where(r =>
                    (r["EstadoNuevo"]?.ToString() ?? "") == "NO PAGADO" &&
                    Convert.ToDateTime(r["FechaAAgregar"]) <= DateTime.Today)
                .OrderBy(r => Convert.ToInt32(r["ID"]))
                .ToList();

            if (!regs.Any()) return;

            int nextRow = hoja.UsedRange.Rows.Count + 1;

            foreach (var row in regs)
            {
                // 🔹 BUSCAR UNA FILA COMPLETAMENTE VACÍA
                while (!FilaVacia(hoja, nextRow))
                {
                    nextRow++;
                }

                // 🔹 PEGAR SI O SI (nunca se saltea)
                for (int col = 1; col <= 22; col++)
                {
                    if (col == 9) continue; // evitar columna I

                    object valor = row[col];

                    // FechaPago → se setea la de hoy al pegar
                    if (col == 3)
                    {
                        // usar SIEMPRE la fecha modelo de la fila 2 columna 3
                        if (fechaPagoModelo.HasValue)
                        {
                            valor = fechaPagoModelo.Value;
                            hoja.Cells[nextRow, col].NumberFormat = "dd/mm/yyyy";
                        }
                        else
                        {
                            valor = "";
                        }
                    }

                    else if (col == 16)
                    {
                        valor = "PENDIENTE-EXCEP ANTICIPO";
                    }

                    hoja.Cells[nextRow, col].Value2 = valor;
                }

                nextRow++;
            }

            // 🔹 MARCAR COMO PAGADO EN BASE
            string ids = string.Join(",", regs.Select(r => r["ID"].ToString()));
            string updFinal = $@"
UPDATE [dbo].[ExcepcionAnticipo]
SET EstadoNuevo = 'PAGADO', 
    FechaPagado = GETDATE(),
    FechaPago = ISNULL(FechaPago, GETDATE())
WHERE ID IN ({ids})
  AND FechaAAgregar <= CAST(GETDATE() AS date);";

            using (var cmd = new SqlCommand(updFinal, conn))
                cmd.ExecuteNonQuery();
        }

        // ---------------------------------------------------------------------
        // 🔹 Utilidades
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
        private bool FilaVacia(Excel.Worksheet hoja, int fila)
        {
            for (int col = 1; col <= 22; col++)
            {
                object v = (hoja.Cells[fila, col] as Excel.Range)?.Value2;
                if (v != null && v.ToString().Trim() != "")
                    return false; // hay datos → NO está vacía
            }
            return true; // TODA vacía → OK
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