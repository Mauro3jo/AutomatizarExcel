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
                            InsertarEnBase(conn, fila, hojaPrincipal, hojaQuitadas, i);

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
                       (cuotas == 1 && totalDescuento > 0.04m) ? 17 : 7;

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
            string fechaPagoTexto = fechaPagoModelo?.ToString("dd/MM/yyyy") ?? "";

            // 🔹 Traer solo registros NO PAGADOS con fechaAAgregar hasta HOY
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

            // 📥 Traer todos los registros a memoria
            DataTable dt = new DataTable();
            using (var adapter = new SqlDataAdapter(sql, conn))
            {
                adapter.Fill(dt);
            }

            // 🔹 Detectar duplicados en memoria (solo primer registro queda NO PAGADO)
            var grupos = dt.AsEnumerable()
                .GroupBy(r => new
                {
                    FechaOperacion = r.Field<DateTime>("FechaOperacion"),
                    NroCupon = r["NroCupon"]?.ToString()?.Trim() ?? "",
                    NroComercio = r["NroComercio"]?.ToString()?.Trim() ?? "",
                    NroTarjeta = r["NroTarjeta"]?.ToString()?.Trim() ?? ""
                });

            foreach (var grupo in grupos)
            {
                bool primero = true;
                foreach (var fila in grupo)
                {
                    string estado = primero ? "NO PAGADO" : "DUPLICADO";
                    fila["EstadoNuevo"] = estado;
                    primero = false;
                }
            }

            // 🔹 Actualizar estados en base (sin tocar pagos futuros)
            foreach (DataRow row in dt.Rows)
            {
                string upd = @"
    UPDATE [dbo].[ExcepcionAnticipo]
    SET EstadoNuevo = @EstadoNuevo,
        FechaPagado = CASE WHEN @EstadoNuevo = 'DUPLICADO' THEN GETDATE() ELSE NULL END
    WHERE ID = @ID;";

                using (var cmd = new SqlCommand(upd, conn))
                {
                    cmd.Parameters.AddWithValue("@EstadoNuevo", row["EstadoNuevo"] ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@ID", row["ID"]);
                    cmd.ExecuteNonQuery();
                }
            }

            // 🔹 Filtrar solo las NO PAGADAS hasta hoy
            var registrosParaPegar = dt.AsEnumerable()
                .Where(r =>
                    (r["EstadoNuevo"]?.ToString() ?? "") == "NO PAGADO" &&
                    Convert.ToDateTime(r["FechaAAgregar"]) <= DateTime.Today)
                .OrderBy(r => Convert.ToInt32(r["ID"]))
                .ToList();

            if (!registrosParaPegar.Any()) return;

            int nextRow = hoja.UsedRange.Rows.Count + 1;

            foreach (var row in registrosParaPegar)
            {
                for (int col = 1; col <= 22; col++)
                {
                    if (col == 9) continue; // ❌ Saltar columna I (anticipo)

                    object valor = row[col];

                    if (col == 3) // Fecha de pago
                    {
                        valor = DateTime.Now.ToString("dd/MM/yyyy");
                        hoja.Cells[nextRow, col].NumberFormat = "@";
                        hoja.Cells[nextRow, col].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        hoja.Cells[nextRow, col].Interior.Color =
                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
                    }
                    else if (col == 16)
                    {
                        valor = "PENDIENTE-EXCEP ANTICIPO";
                    }

                    hoja.Cells[nextRow, col].Value2 = valor;
                }
                nextRow++;
            }

            // 🔹 Actualizar solo las filas agregadas a PAGADO
            string ids = string.Join(",", registrosParaPegar.Select(r => r["ID"].ToString()));

            string updFinal = $@"
    UPDATE [dbo].[ExcepcionAnticipo]
    SET EstadoNuevo = 'PAGADO', 
        FechaPagado = GETDATE(),
        FechaPago = ISNULL(FechaPago, GETDATE())
    WHERE ID IN ({ids})
      AND FechaAAgregar <= CAST(GETDATE() AS date);";

            using (var cmd = new SqlCommand(updFinal, conn))
            {
                cmd.ExecuteNonQuery();
            }
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
