using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;

namespace Automatizacion_excel.Paso2
{
    public class SubirExcelAnticipo
    {
        private readonly string connectionString;

        public SubirExcelAnticipo()
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            connectionString = config.GetConnectionString("MiConexion");
        }

        public void Ejecutar()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";
            ofd.Title = "Seleccioná el Excel de anticipos";

            if (ofd.ShowDialog() != DialogResult.OK)
                return;

            try
            {
                using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                {
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        var table = result.Tables[0];

                        int insertados = 0, ignorados = 0;

                        using (var conn = new SqlConnection(connectionString))
                        {
                            conn.Open();

                            for (int i = 1; i < table.Rows.Count; i++)
                            {
                                DataRow row = table.Rows[i];
                                string tarjeta = row[18]?.ToString()?.Trim() ?? "";

                                if (string.IsNullOrWhiteSpace(tarjeta) ||
                                    tarjeta.Contains("QR", StringComparison.OrdinalIgnoreCase) ||
                                    tarjeta.Contains("REINTEGRO", StringComparison.OrdinalIgnoreCase))
                                {
                                    ignorados++;
                                    continue;
                                }

                                // cuotas
                                int.TryParse(row[16]?.ToString(), out int cuotas);
                                string tipoPago = cuotas == 0 ? "DÉBITO" :
                                                  cuotas == 1 ? "CRÉDITO 1 PAGO" :
                                                  "CRÉDITO 2 O MÁS PAGOS";

                                // total descuento
                                string descuentoStr = row[8]?.ToString()?.Replace("%", "").Replace(",", ".").Trim();
                                double.TryParse(descuentoStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double descuento);

                                // detectar tipo de tarjeta
                                string categoriaTarjeta = DetectarCategoriaTarjeta(tarjeta, tipoPago, descuento);

                                // buscar días en PlazosDeAcreditaciones
                                int dias = ObtenerDiasPlazo(conn, tipoPago, categoriaTarjeta);

                                // fechas
                                DateTime fechaOperacion = ParsearFecha(row[0]);
                                DateTime fechaPago = ParsearFecha(row[2]);
                                DateTime fechaAAgregar = SumarDiasHabiles(fechaOperacion, dias);

                                // insertar
                                string sql = @"
                                    INSERT INTO [dbo].[ExcepcionAnticipo]
                                    ([FechaOperacion],[FechaPresentacion],[FechaPago],[NroCupon],
                                     [NroComercio],[NroTarjeta],[Moneda],[TotalBruto],[TotalDescuento],[TotalNeto],
                                     [EntidadPagadora],[CuentaBancaria],[NroLiquidacion],[NroLote],[TipoLiquidacion],
                                     [Estado],[Cuotas],[NroAutorizacion],[Tarjeta],[TipoOperacion],[ComercioParticipante],
                                     [PromocionPlan],[DiasAdelanto],[FechaAAgregar],[EstadoNuevo],[PagoOriginal],[FechaPagado])
                                    VALUES
                                    (@FechaOperacion,@FechaPresentacion,@FechaPago,@NroCupon,
                                     @NroComercio,@NroTarjeta,@Moneda,@TotalBruto,@TotalDescuento,@TotalNeto,
                                     @EntidadPagadora,@CuentaBancaria,@NroLiquidacion,@NroLote,@TipoLiquidacion,
                                     @Estado,@Cuotas,@NroAutorizacion,@Tarjeta,@TipoOperacion,@ComercioParticipante,
                                     @PromocionPlan,@DiasAdelanto,@FechaAAgregar,@EstadoNuevo,@PagoOriginal,@FechaPagado);";

                                using (var cmd = new SqlCommand(sql, conn))
                                {
                                    cmd.Parameters.AddWithValue("@FechaOperacion", fechaOperacion);
                                    cmd.Parameters.AddWithValue("@FechaPresentacion", row[1] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@FechaPago", fechaPago);
                                    cmd.Parameters.AddWithValue("@NroCupon", row[3] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@NroComercio", row[4] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@NroTarjeta", row[5] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Moneda", row[6] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@TotalBruto", LimpiarDecimal(row[7]));
                                    cmd.Parameters.AddWithValue("@TotalDescuento", LimpiarDecimal(row[8]));
                                    cmd.Parameters.AddWithValue("@TotalNeto", LimpiarDecimal(row[9]));
                                    cmd.Parameters.AddWithValue("@EntidadPagadora", row[10] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@CuentaBancaria", row[11] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@NroLiquidacion", row[12] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@NroLote", row[13] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@TipoLiquidacion", row[14] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Estado", row[15] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Cuotas", cuotas);
                                    cmd.Parameters.AddWithValue("@NroAutorizacion", row[17] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@Tarjeta", tarjeta);
                                    cmd.Parameters.AddWithValue("@TipoOperacion", row[19] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@ComercioParticipante", row[20] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@PromocionPlan", row[21] ?? (object)DBNull.Value);
                                    cmd.Parameters.AddWithValue("@DiasAdelanto", dias);
                                    cmd.Parameters.AddWithValue("@FechaAAgregar", fechaAAgregar);
                                    cmd.Parameters.AddWithValue("@EstadoNuevo", "NO PAGADO");
                                    cmd.Parameters.AddWithValue("@PagoOriginal", fechaPago);
                                    cmd.Parameters.AddWithValue("@FechaPagado", DBNull.Value);

                                    cmd.ExecuteNonQuery();
                                    insertados++;
                                }
                            }
                        }

                        MessageBox.Show($"Carga completada.\nInsertados: {insertados}\nIgnorados: {ignorados}", "Éxito");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al procesar el archivo:\n" + ex.Message, "Error");
            }
        }

        private static string DetectarCategoriaTarjeta(string tarjeta, string tipoPago, double descuento)
        {
            string categoria = "";

            tarjeta = tarjeta.ToUpper();

            if (tarjeta.Contains("CABAL")) categoria = "Bancarizadas (Cabal)";
            else if (tarjeta.Contains("AMEX")) categoria = "Bancarizadas (Amex)";
            else if (tarjeta.Contains("VISA") || tarjeta.Contains("MASTER") || tarjeta.Contains("ARGENCARD"))
                categoria = "Bancarizadas (Visa - Master - ArgenCard)";
            else if (tarjeta.Contains("NARANJA")) categoria = "No Bancarizadas (Naranja Visa, Naranja Master, Cencosud, etc.)";

            // Si es crédito 1 pago y descuento alto => recargable
            if (tipoPago == "CRÉDITO 1 PAGO" && descuento >= 6)
                categoria = "Recargables (Tarjetas de débito de bancos virtuales, Ualá, Brubank, Wilobank, ReBa, Banco del Sol, MercadoPago y todas las recargables)";

            return categoria;
        }

        private static DateTime ParsearFecha(object valor)
        {
            if (valor == null || valor == DBNull.Value)
                return DateTime.MinValue;

            if (DateTime.TryParse(valor.ToString(), out var fecha))
                return fecha;

            if (double.TryParse(valor.ToString(), out var oa))
                return DateTime.FromOADate(oa);

            return DateTime.MinValue;
        }

        private int ObtenerDiasPlazo(SqlConnection conn, string tipoPago, string categoriaTarjeta)
        {
            string sql = @"SELECT TOP 1 Dias 
                           FROM [dbo].[PlazosDeAcreditaciones]
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
            if (fechaInicio == DateTime.MinValue || dias <= 0)
                return fechaInicio;

            int agregados = 0;
            DateTime fecha = fechaInicio;

            while (agregados < dias)
            {
                fecha = fecha.AddDays(1);
                if (fecha.DayOfWeek != DayOfWeek.Saturday && fecha.DayOfWeek != DayOfWeek.Sunday)
                    agregados++;
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
