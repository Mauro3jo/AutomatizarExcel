using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace Automatizacion_excel.RecortarExcel
{
    internal class Recortar_Excel2
    {
        internal static bool ProcesarArchivo(string archivoEntrada, string archivoSalida)
        {
            int movimientosCopiados = 0;

            try
            {
                using (var workbook = new XLWorkbook(archivoEntrada))
                {
                    var wsOrigen = workbook.Worksheet(1);
                    var wbDestino = new XLWorkbook();
                    var wsDestino = wbDestino.AddWorksheet("Movimientos");

                    // 1. Buscar fila de encabezados y mapear las posiciones
                    int filaCabecera = -1;
                    int ultimaFila = wsOrigen.LastRowUsed().RowNumber();

                    // Diccionarios para posiciones de cada encabezado
                    var encabezados = new Dictionary<string, int>();
                    var columnasMovimiento = new List<int>();

                    for (int f = 1; f <= ultimaFila; f++)
                    {
                        int ultimaCol = wsOrigen.Row(f).LastCellUsed()?.Address.ColumnNumber ?? 0;
                        for (int c = 1; c <= ultimaCol; c++)
                        {
                            var valor = wsOrigen.Cell(f, c).GetString().Trim().ToUpper();

                            // Ajusta según tus posibles encabezados (agrega variantes si necesitas)
                            if (valor.Contains("FECHA") && !encabezados.ContainsKey("FECHA"))
                                encabezados["FECHA"] = c;
                            if (valor.Contains("COMPROBANTE") && !encabezados.ContainsKey("COMPROBANTE"))
                                encabezados["COMPROBANTE"] = c;
                            if (valor.Contains("MOVIMIENTO"))
                                columnasMovimiento.Add(c); // Puede haber más de una columna de Movimiento
                            if ((valor.Contains("DÉBITO") || valor.Contains("DEBITO")) && !encabezados.ContainsKey("DÉBITO"))
                                encabezados["DÉBITO"] = c;
                            if ((valor.Contains("CRÉDITO") || valor.Contains("CREDITO")) && !encabezados.ContainsKey("CRÉDITO"))
                                encabezados["CRÉDITO"] = c;
                            if (valor.Contains("SALDO") && !encabezados.ContainsKey("SALDO"))
                                encabezados["SALDO"] = c;
                        }

                        // Consideramos fila de encabezado válida si tiene al menos FECHA, COMPROBANTE, MOVIMIENTO
                        if (encabezados.ContainsKey("FECHA") && encabezados.ContainsKey("COMPROBANTE") && columnasMovimiento.Count > 0)
                        {
                            filaCabecera = f;
                            break;
                        }
                        // Limpiar por si no era la fila de encabezado y seguimos buscando
                        encabezados.Clear();
                        columnasMovimiento.Clear();
                    }

                    if (filaCabecera == -1)
                        return false; // No encontró la tabla

                    // 2. Copiar cabecera destino
                    wsDestino.Cell(1, 1).Value = "Fecha";
                    wsDestino.Cell(1, 2).Value = "Comprobante";
                    wsDestino.Cell(1, 3).Value = "Movimiento";
                    wsDestino.Cell(1, 4).Value = "Débito";
                    wsDestino.Cell(1, 5).Value = "Crédito";
                    wsDestino.Cell(1, 6).Value = "Saldo en cuenta";

                    int filaDestino = 2;
                    string fechaActual = "";

                    // 3. Procesar filas de datos usando los encabezados detectados
                    for (int f = filaCabecera + 1; f <= ultimaFila; f++)
                    {
                        // Si toda la fila está vacía, cortar
                        bool vacia = true;
                        int ultimaCol = wsOrigen.Row(f).LastCellUsed()?.Address.ColumnNumber ?? 0;
                        for (int c = 1; c <= ultimaCol; c++)
                        {
                            if (!string.IsNullOrWhiteSpace(wsOrigen.Cell(f, c).GetString()))
                            {
                                vacia = false;
                                break;
                            }
                        }
                        if (vacia)
                            break;

                        // 1. Fecha (relleno hacia abajo si está vacía)
                        string fecha = "";
                        if (encabezados.ContainsKey("FECHA"))
                        {
                            fecha = wsOrigen.Cell(f, encabezados["FECHA"]).GetString().Trim();
                        }
                        if (!string.IsNullOrEmpty(fecha))
                            fechaActual = fecha;
                        else
                            fecha = fechaActual;

                        // 2. Comprobante
                        string comprobante = encabezados.ContainsKey("COMPROBANTE")
                            ? wsOrigen.Cell(f, encabezados["COMPROBANTE"]).GetString().Trim()
                            : "";

                        // 3. Movimiento (concatenar si hay varias)
                        string movimiento = "";
                        foreach (var col in columnasMovimiento)
                        {
                            var val = wsOrigen.Cell(f, col).GetString().Trim();
                            if (!string.IsNullOrEmpty(val))
                            {
                                if (movimiento.Length > 0) movimiento += " ";
                                movimiento += val;
                            }
                        }

                        // 4. Débito
                        string debito = encabezados.ContainsKey("DÉBITO")
                            ? wsOrigen.Cell(f, encabezados["DÉBITO"]).GetString().Trim()
                            : "";

                        // 5. Crédito
                        string credito = encabezados.ContainsKey("CRÉDITO")
                            ? wsOrigen.Cell(f, encabezados["CRÉDITO"]).GetString().Trim()
                            : "";

                        // 6. Saldo
                        string saldo = encabezados.ContainsKey("SALDO")
                            ? wsOrigen.Cell(f, encabezados["SALDO"]).GetString().Trim()
                            : "";

                        // Ignorar si todo vacío
                        if (string.IsNullOrWhiteSpace(comprobante) &&
                            string.IsNullOrWhiteSpace(movimiento) &&
                            string.IsNullOrWhiteSpace(debito) &&
                            string.IsNullOrWhiteSpace(credito) &&
                            string.IsNullOrWhiteSpace(saldo))
                            continue;

                        // Copiar a destino
                        wsDestino.Cell(filaDestino, 1).Value = fecha;
                        wsDestino.Cell(filaDestino, 2).Value = comprobante;
                        wsDestino.Cell(filaDestino, 3).Value = movimiento;
                        wsDestino.Cell(filaDestino, 4).Value = debito;
                        wsDestino.Cell(filaDestino, 5).Value = credito;
                        wsDestino.Cell(filaDestino, 6).Value = saldo;

                        filaDestino++;
                        movimientosCopiados++;
                    }

                    wbDestino.SaveAs(archivoSalida);
                }

                return movimientosCopiados > 0;
            }
            catch
            {
                throw;
            }
        }
    }
}
