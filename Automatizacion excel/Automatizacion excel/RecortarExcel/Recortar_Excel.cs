using System;
using ClosedXML.Excel;

namespace Automatizacion_excel.RecortarExcel
{
    public static class Recortar_Excel
    {
        public static bool ProcesarArchivo(string archivoEntrada, string archivoSalida)
        {
            int movimientosCopiados = 0;
            try
            {
                using (var workbook = new XLWorkbook(archivoEntrada))
                {
                    var wsOrigen = workbook.Worksheet(1);
                    var wbDestino = new XLWorkbook();
                    var wsDestino = wbDestino.AddWorksheet("Movimientos");

                    int filaDestino = 1;
                    int ultimaFila = wsOrigen.LastRowUsed().RowNumber();
                    bool dentroDeBloque = false;
                    string[] ultimoMovimiento = new string[14];

                    int idxDebito = -1, idxCredito = -1, idxSaldo = -1;
                    bool indicesCabeceraListos = false;

                    // Flag para esperar destinatario
                    bool esperarDestinatario = false;
                    int filaDestinoEsperada = 0;

                    for (int f = 1; f <= ultimaFila; f++)
                    {
                        // Detectar cabecera en C a P
                        bool tieneFecha = false, tieneConcepto = false, tieneComprobante = false, tieneDebito = false, tieneCredito = false, tieneSaldo = false;
                        for (int c = 3; c <= 16; c++)
                        {
                            var val = wsOrigen.Cell(f, c).GetString().Trim().ToUpper();
                            if (val.Contains("FECHA")) tieneFecha = true;
                            if (val.Contains("CONCEPTO")) tieneConcepto = true;
                            if (val.Contains("COMPROBANTE")) tieneComprobante = true;
                            if (val.Contains("DEBITO")) tieneDebito = true;
                            if (val.Contains("CREDITO")) tieneCredito = true;
                            if (val.Contains("SALDO")) tieneSaldo = true;
                        }
                        bool esCabecera = tieneFecha && tieneConcepto && tieneComprobante && tieneDebito && tieneCredito && tieneSaldo;

                        if (esCabecera)
                        {
                            dentroDeBloque = true;
                            if (filaDestino == 1)
                            {
                                for (int c = 3; c <= 16; c++)
                                    wsDestino.Cell(filaDestino, c - 2).Value = wsOrigen.Cell(f, c).GetString();
                                filaDestino++;
                            }

                            // Detectar los índices una sola vez
                            if (!indicesCabeceraListos)
                            {
                                for (int c = 3; c <= 16; c++)
                                {
                                    var val = wsOrigen.Cell(f, c).GetString().Trim().ToUpper();
                                    if (val.Contains("DEBITO")) idxDebito = c - 3;
                                    if (val.Contains("CREDITO")) idxCredito = c - 3;
                                    if (val.Contains("SALDO")) idxSaldo = c - 3;
                                }
                                indicesCabeceraListos = true;
                            }
                            continue;
                        }

                        if (!dentroDeBloque)
                            continue;

                        // Si la fila está completamente vacía, saltar
                        bool vacia = true;
                        for (int c = 3; c <= 16; c++)
                        {
                            if (!string.IsNullOrWhiteSpace(wsOrigen.Cell(f, c).GetString()))
                            {
                                vacia = false;
                                break;
                            }
                        }
                        if (vacia) continue;

                        // Detectar corte por texto especial
                        string textoFilaCompleta = "";
                        for (int c = 3; c <= 16; c++)
                            textoFilaCompleta += wsOrigen.Cell(f, c).GetString().ToUpper() + " ";
                        if (textoFilaCompleta.Contains("ZOCO SAS") || textoFilaCompleta.Contains("ESTADO DE CUENTAS"))
                        {
                            // ----- EXCEPCIÓN: Si la fila anterior fue transferencia/debito y la siguiente solo tiene destinatario, concatenar ----
                            bool ultimaFueOperacionEspecial = false;
                            for (int cc = 3; cc <= 9; cc++)
                            {
                                string val = wsOrigen.Cell(f - 1, cc).GetString().ToUpper();
                                if (EsOperacionEspecial(val))
                                {
                                    ultimaFueOperacionEspecial = true;
                                    break;
                                }
                            }
                            if (ultimaFueOperacionEspecial && f < ultimaFila)
                            {
                                esperarDestinatario = true;
                                filaDestinoEsperada = filaDestino - 1;
                            }
                            dentroDeBloque = false;
                            continue;
                        }

                        // Unir destinatario si flag activo
                        if (esperarDestinatario)
                        {
                            bool soloDestinatario = true;
                            string dest = "";
                            for (int c = 10; c <= 16; c++)
                            {
                                if (!string.IsNullOrWhiteSpace(wsOrigen.Cell(f, c).GetString()))
                                {
                                    soloDestinatario = false;
                                    break;
                                }
                            }
                            for (int c = 3; c <= 9; c++)
                                dest += wsOrigen.Cell(f, c).GetString().Trim() + " ";
                            dest = dest.Trim();

                            if (soloDestinatario && !string.IsNullOrWhiteSpace(dest) && filaDestinoEsperada > 0)
                            {
                                var celdaActual = wsDestino.Cell(filaDestinoEsperada, 2);
                                string textoActual = celdaActual.GetString();
                                wsDestino.Cell(filaDestinoEsperada, 2).Value =
                                    (!string.IsNullOrWhiteSpace(textoActual) ? textoActual + " " : "") + dest;
                                esperarDestinatario = false;
                                continue; // No grabar esta fila como movimiento
                            }
                        }

                        // Ignorar "CTA CORR", "SALDO RESUMEN"
                        var vC = wsOrigen.Cell(f, 3).GetString().Trim().ToUpper();
                        if (vC.Contains("CTA CORR") || vC.Contains("SALDO RESUMEN"))
                            continue;

                        // COPIA LITERAL DE C a P (3 a 16)
                        for (int c = 3; c <= 16; c++)
                            ultimoMovimiento[c - 3] = wsOrigen.Cell(f, c).GetString();

                        // Revisar si en alguna celda de C a I hay operación especial
                        bool tieneUnionEspecial = false;
                        for (int c = 3; c <= 9; c++)
                        {
                            string valor = wsOrigen.Cell(f, c).GetString().ToUpper();
                            if (EsOperacionEspecial(valor))
                            {
                                tieneUnionEspecial = true;
                                break;
                            }
                        }

                        if (tieneUnionEspecial && (f < ultimaFila))
                        {
                            bool siguienteSoloCI = true;
                            for (int c = 10; c <= 16; c++)
                            {
                                if (!string.IsNullOrWhiteSpace(wsOrigen.Cell(f + 1, c).GetString()))
                                {
                                    siguienteSoloCI = false;
                                    break;
                                }
                            }

                            string textoSiguienteCI = "";
                            for (int c = 3; c <= 9; c++)
                            {
                                var val = wsOrigen.Cell(f + 1, c).GetString().Trim();
                                if (!string.IsNullOrWhiteSpace(val))
                                {
                                    if (textoSiguienteCI.Length > 0) textoSiguienteCI += " ";
                                    textoSiguienteCI += val;
                                }
                            }

                            if (siguienteSoloCI && !string.IsNullOrWhiteSpace(textoSiguienteCI))
                            {
                                if (!string.IsNullOrWhiteSpace(ultimoMovimiento[1]))
                                    ultimoMovimiento[1] += " " + textoSiguienteCI;
                                else
                                    ultimoMovimiento[1] = textoSiguienteCI;

                                f++; // Salteamos la fila siguiente porque ya la usamos
                            }
                        }

                        // ------ AJUSTE DEBITO/CREDITO/SALDO: SPLIT SOLO SI CORRESPONDE ------
                        if (indicesCabeceraListos && idxDebito != -1 && idxCredito != -1 && idxSaldo != -1)
                        {
                            string valDebito = ultimoMovimiento[idxDebito]?.Trim() ?? "";
                            string valCredito = ultimoMovimiento[idxCredito]?.Trim() ?? "";
                            string valSaldo = ultimoMovimiento[idxSaldo]?.Trim() ?? "";

                            var partesDebito = valDebito.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            var partesCredito = valCredito.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                            // Split en Débito si corresponde
                            if (partesDebito.Length == 2 && string.IsNullOrWhiteSpace(valCredito) && string.IsNullOrWhiteSpace(valSaldo))
                            {
                                ultimoMovimiento[idxDebito] = "";
                                ultimoMovimiento[idxCredito] = partesDebito[0];
                                ultimoMovimiento[idxSaldo] = partesDebito[1];
                            }
                            // Split en Crédito si corresponde
                            else if (partesCredito.Length == 2 && string.IsNullOrWhiteSpace(valDebito) && string.IsNullOrWhiteSpace(valSaldo))
                            {
                                ultimoMovimiento[idxDebito] = "";
                                ultimoMovimiento[idxCredito] = partesCredito[0];
                                ultimoMovimiento[idxSaldo] = partesCredito[1];
                            }
                            // Si hay uno en Débito y uno en Crédito
                            else if (partesDebito.Length == 1 && !string.IsNullOrWhiteSpace(valCredito))
                            {
                                ultimoMovimiento[idxDebito] = partesDebito[0];
                                ultimoMovimiento[idxSaldo] = valCredito;
                                ultimoMovimiento[idxCredito] = "";
                            }
                            // Si hay uno en Crédito y uno en Débito (caso inverso)
                            else if (partesCredito.Length == 1 && !string.IsNullOrWhiteSpace(valDebito))
                            {
                                ultimoMovimiento[idxCredito] = partesCredito[0];
                                ultimoMovimiento[idxSaldo] = valDebito;
                                ultimoMovimiento[idxDebito] = "";
                            }
                            // Los demás casos quedan igual
                        }

                        // Guardar el movimiento resultante en la hoja destino
                        for (int i = 0; i < 14; i++)
                            wsDestino.Cell(filaDestino, i + 1).Value = ultimoMovimiento[i] ?? "";

                        filaDestino++;
                        movimientosCopiados++; // <<<--- Cuenta solo filas útiles
                    }

                    wbDestino.SaveAs(archivoSalida);
                }

                // Devuelve si se copiaron movimientos
                return movimientosCopiados > 0;
            }
            catch
            {
                throw; // Dejá que el formulario lo maneje
            }
        }

        private static bool EsOperacionEspecial(string texto)
        {
            texto = texto.ToUpper();
            return texto.Contains("TRANSFERENCIA")
                || texto.Contains("DEBITO")
                || texto.Contains("PAGO A PROVEEDORES")
                || texto.Contains("TEF DATANET");
        }
    }
}
