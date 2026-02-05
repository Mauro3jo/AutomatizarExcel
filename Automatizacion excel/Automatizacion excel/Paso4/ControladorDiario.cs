using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static Automatizacion_excel.Paso4.ControladorDiario;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatizacion_excel.Paso4
{
    public class ControladorDiario
    {
        private readonly string rutaArchivo;

        public ControladorDiario(string ruta)
        {
            rutaArchivo = ruta;
        }


        public string ControlarFechaUnica(out List<int> filasInvalidas, Action<string, int> reportar)
        {
            filasInvalidas = new List<int>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                reportar("📂 Abriendo archivo...", 5);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hoja = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;

                if (hoja == null)
                    throw new Exception("No se encontró la hoja 'Reporte Diario2'.");

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                string fechaUnica = null;
                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    var celda = hoja.Cells[fila, 3] as Excel.Range; // Columna C
                    string valor = celda?.Text?.Trim();

                    if (string.IsNullOrEmpty(valor)) continue;

                    if (fechaUnica == null)
                        fechaUnica = valor;
                    else if (valor != fechaUnica)
                        filasInvalidas.Add(fila);

                    int progreso = (int)((fila - 1f) / (ultimaFila - 1f) * 40);
                    reportar($"📅 Verificando fecha en fila {fila}...", progreso);
                }

                if (filasInvalidas.Count == 0)
                    return $"✅ Fecha única detectada: {fechaUnica}";
                else
                    return $"❌ Fechas diferentes detectadas en filas: {string.Join(", ", filasInvalidas)}";
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }
        public (string, string) ValidarArancelEIVA(Action<string, int> reportar)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            int COL_BRUTO = 8;       // H
            int COL_COMISION = 28;   // AB
            int COL_ARANCEL = 29;    // AC
            int COL_IVA = 30;        // AD

            List<int> filasArancelMal = new List<int>();
            List<int> filasIvaMal = new List<int>();

            try
            {
                reportar("🧮 Validando Arancel e IVA...", 60);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hoja = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;

                if (hoja == null)
                    throw new Exception("No se encontró la hoja 'Reporte Diario2'.");

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    // ===== BRUTO =====
                    string brutoTxt = (hoja.Cells[fila, COL_BRUTO] as Excel.Range)?.Value2?.ToString();
                    double bruto = 0;

                    if (!string.IsNullOrWhiteSpace(brutoTxt))
                    {
                        brutoTxt = brutoTxt
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Replace("(", "-")
                            .Replace(")", "")
                            .Replace(".", "")
                            .Replace(",", ".")
                            .Trim();

                        double.TryParse(
                            brutoTxt,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out bruto
                        );
                    }

                    if (bruto <= 0)
                        continue;

                    // ===== % COMISIÓN =====
                    string comisionTxt = (hoja.Cells[fila, COL_COMISION] as Excel.Range)?.Text
                        ?.Replace("%", "")
                        ?.Replace(",", ".")
                        ?.Trim();

                    double porcentajeComision = 0;
                    double.TryParse(
                        comisionTxt,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out porcentajeComision
                    );

                    double arancelEsperado = Math.Round(bruto * (porcentajeComision / 100), 2);

                    // ===== ARANCEL HOJA =====
                    string arancelTxt = (hoja.Cells[fila, COL_ARANCEL] as Excel.Range)?.Value2?.ToString()
                        ?.Replace("$", "")
                        ?.Replace(".", "")
                        ?.Replace(",", ".")
                        ?.Trim();

                    double arancelHoja = 0;
                    double.TryParse(
                        arancelTxt,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out arancelHoja
                    );

                    if (Math.Abs(arancelHoja - arancelEsperado) > 0.01)
                        filasArancelMal.Add(fila);

                    // ===== IVA =====
                    double ivaEsperado = Math.Round(arancelEsperado * 0.21, 2);

                    string ivaTxt = (hoja.Cells[fila, COL_IVA] as Excel.Range)?.Value2?.ToString()
                        ?.Replace("$", "")
                        ?.Replace(".", "")
                        ?.Replace(",", ".")
                        ?.Trim();

                    double ivaHoja = 0;
                    double.TryParse(
                        ivaTxt,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out ivaHoja
                    );

                    if (Math.Abs(ivaHoja - ivaEsperado) > 0.01)
                        filasIvaMal.Add(fila);

                    int progreso = 60 + (int)((fila - 1f) / (ultimaFila - 1f) * 30);
                    reportar($"🧮 Validando arancel e IVA fila {fila}...", progreso);
                }

                string resultadoArancel = filasArancelMal.Count == 0
                    ? "✅ Todos los aranceles están correctos."
                    : $"❌ Error en arancel en filas: {string.Join(", ", filasArancelMal)}";

                string resultadoIva = filasIvaMal.Count == 0
                    ? "✅ Todos los IVA 21% están correctos."
                    : $"❌ Error en IVA en filas: {string.Join(", ", filasIvaMal)}";

                return (resultadoArancel, resultadoIva);
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        public string ValidarFUR(Action<string, int> reportar)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            try
            {
                reportar("🧐 Validando FUR...", 0);
                excelApp = new Excel.Application { DisplayAlerts = false };
                workbook = excelApp.Workbooks.Open(rutaArchivo);

                Excel.Worksheet hojaFUR = workbook.Sheets["FUR"] as Excel.Worksheet;
                Excel.Worksheet hojaDiario = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;
                if (hojaFUR == null || hojaDiario == null)
                    throw new Exception("No se encontró la hoja FUR o Reporte Diario2.");

                var datosFUR = hojaFUR.UsedRange.Value2 as object[,];
                var datosDiario = hojaDiario.UsedRange.Value2 as object[,];

                int offsetFURRow = hojaFUR.UsedRange.Row;
                int offsetFURCol = hojaFUR.UsedRange.Column;
                int offsetDiarioRow = hojaDiario.UsedRange.Row;
                int offsetDiarioCol = hojaDiario.UsedRange.Column;

                int filasFUR = datosFUR.GetLength(0);
                int filasDiario = datosDiario.GetLength(0);

                int filaDatosFUR = 10;
                int colRazon = 2, colComercio = 3, colCBU = 4, colCuenta = 5, colLegajo = 6, colCUIT = 7, colCP = 8, colImporte = 9;
                int colPagoReal = 57, colRazonD = 52, colComercioD = 44, colCBUD = 39, colCuentaD = 42, colLegajoD = 53, colCUITD = 32, colCPD = 54, colTotalDesc = 38;

                // Armar lista de pagos del Diario
                var listaDiario = new List<(string razon, string comercio, string cbu, string cuenta, string legajo, string cuit, string cp, double totalDesc)>();
                for (int fila = offsetDiarioRow + 1; fila <= filasDiario; fila++)
                {
                    int r = fila - offsetDiarioRow + 1;
                    string pagoReal = Safe(datosDiario[r, colPagoReal - offsetDiarioCol + 1]);
                    if (pagoReal != "SI PAGA") continue;

                    string razon = Safe(datosDiario[r, colRazonD - offsetDiarioCol + 1]);
                    string comercio = Safe(datosDiario[r, colComercioD - offsetDiarioCol + 1]);
                    string cbu = Safe(datosDiario[r, colCBUD - offsetDiarioCol + 1]);
                    string cuenta = Safe(datosDiario[r, colCuentaD - offsetDiarioCol + 1]);
                    string legajo = Safe(datosDiario[r, colLegajoD - offsetDiarioCol + 1]);
                    string cuit = Safe(datosDiario[r, colCUITD - offsetDiarioCol + 1]);
                    string cp = Safe(datosDiario[r, colCPD - offsetDiarioCol + 1]);

                    double totalDesc = ParseImporte(datosDiario[r, colTotalDesc - offsetDiarioCol + 1]);

                    listaDiario.Add((razon, comercio, cbu, cuenta, legajo, cuit, cp, totalDesc));
                }

                List<string> errores = new();

                for (int filaFUR = filaDatosFUR; filaFUR <= filasFUR; filaFUR++)
                {
                    int f = filaFUR - offsetFURRow + 1;
                    string razonF = Safe(datosFUR[f, colRazon - offsetFURCol + 1]);

                    if (razonF == "TOTAL GENERAL")
                    {
                        double sumaTotal = listaDiario.Sum(x => x.totalDesc);
                        double importeTotalGeneral = ParseImporte(datosFUR[f, colImporte - offsetFURCol + 1]);
                        if (Math.Abs(sumaTotal - importeTotalGeneral) > 0.01)
                            errores.Add($"❌ Fila {filaFUR} (FUR): El TOTAL GENERAL esperado es {sumaTotal:N2} pero figura {importeTotalGeneral:N2}.");
                        continue;
                    }

                    string comercioF = Safe(datosFUR[f, colComercio - offsetFURCol + 1]);
                    string cbuF = Safe(datosFUR[f, colCBU - offsetFURCol + 1]);
                    string cuentaF = Safe(datosFUR[f, colCuenta - offsetFURCol + 1]);
                    string legajoF = Safe(datosFUR[f, colLegajo - offsetFURCol + 1]);
                    string cuitF = Safe(datosFUR[f, colCUIT - offsetFURCol + 1]);
                    string cpF = Safe(datosFUR[f, colCP - offsetFURCol + 1]);
                    double importeF = ParseImporte(datosFUR[f, colImporte - offsetFURCol + 1]);

                    // Buscar todas las filas FUR que coinciden exactamente con estos datos (puede haber varias)
                    var filasCoincidentesFUR = new List<int>();
                    double sumaFUR = 0;
                    for (int j = filaDatosFUR; j <= filasFUR; j++)
                    {
                        int idx = j - offsetFURRow + 1;
                        string razonX = Safe(datosFUR[idx, colRazon - offsetFURCol + 1]);
                        string comercioX = Safe(datosFUR[idx, colComercio - offsetFURCol + 1]);
                        string cbuX = Safe(datosFUR[idx, colCBU - offsetFURCol + 1]);
                        string cuentaX = Safe(datosFUR[idx, colCuenta - offsetFURCol + 1]);
                        string legajoX = Safe(datosFUR[idx, colLegajo - offsetFURCol + 1]);
                        string cuitX = Safe(datosFUR[idx, colCUIT - offsetFURCol + 1]);
                        string cpX = Safe(datosFUR[idx, colCP - offsetFURCol + 1]);

                        if (razonX == razonF && comercioX == comercioF && cbuX == cbuF && cuentaX == cuentaF && legajoX == legajoF && cuitX == cuitF && cpX == cpF)
                        {
                            double imp = ParseImporte(datosFUR[idx, colImporte - offsetFURCol + 1]);
                            sumaFUR += imp;
                            filasCoincidentesFUR.Add(j);
                        }
                    }

                    // Buscar todas las filas en Diario que coinciden con estos datos (puede+ haber varias)
                    var coincidenciasDiario = listaDiario.Where(x =>
                        x.razon == razonF && x.comercio == comercioF && x.cbu == cbuF &&
                        x.cuenta == cuentaF && x.legajo == legajoF && x.cuit == cuitF && x.cp == cpF
                    ).ToList();
                    double sumaDiario = coincidenciasDiario.Sum(x => x.totalDesc);

                    // Si la suma de los importes de las filas FUR coincidentes da igual al Diario, es válido
                    if (Math.Abs(sumaDiario - sumaFUR) > 0.01)
                    {
                        // Si no, error, pero sigue con los otros
                        errores.Add($"❌ Fila(s) {string.Join(", ", filasCoincidentesFUR)} (FUR): Importe(s) FUR suman {sumaFUR:N2} pero en Diario suman {sumaDiario:N2}.");
                    }

                    int prog = (int)(100.0 * (filaFUR - filaDatosFUR) / (filasFUR - filaDatosFUR + 1));
                    reportar($"🧐 Validando FUR fila {filaFUR}...", prog);
                }

                var faltantes = ObtenerFaltantesEnFUR((msg, prog) => { });
                if (faltantes.Count > 0)
                {
                    var formFaltantes = new FaltantesForm(faltantes);
                    formFaltantes.ShowDialog();
                }

                return errores.Count == 0
                    ? "✅ FUR coincide perfectamente con los datos del Reporte Diario2."
                    : "❌ Errores detectados en FUR:\n" + string.Join("\n", errores);
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private string Safe(object value) => Convert.ToString(value ?? "").Trim().ToUpper();

        private double ParseImporte(object value)
        {
            string texto = Convert.ToString(value ?? "0").Replace("$", "").Replace(".", "").Replace(",", ".").Trim();
            double.TryParse(texto, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double r);
            return r;
        }

        private bool Filtrar(ref List<string> posibles, int index, string esperado, int fila, List<string> errores, string campo)
        {
            posibles = posibles.Where(k => k.Split('|')[index] == esperado).ToList();
            if (posibles.Count == 0)
            {
                errores.Add($"❌ Fila {fila} (FUR): No se encontró {campo} '{esperado}'.");
                return false;
            }
            return true;
        }


        public string SumarColumnaBruto(Action<string, int> reportar)
        {
            double suma = 0;
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                reportar("➕ Iniciando suma de BRUTO...", 50);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hoja = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;

                if (hoja == null)
                    throw new Exception("No se encontró la hoja 'Reporte Diario2'.");

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    var celda = hoja.Cells[fila, 8] as Excel.Range; // Columna H (BRUTO)

                    // ✅ leer y parsear sin ambigüedad + soporta $ ( ) , .
                    string txt = celda?.Value2?.ToString();
                    if (!string.IsNullOrWhiteSpace(txt))
                    {
                        txt = txt
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Replace("(", "-")
                            .Replace(")", "")
                            .Replace(".", "")   // miles
                            .Replace(",", ".")  // decimal
                            .Trim();

                        if (double.TryParse(
                                txt,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out double valor
                            ))
                        {
                            suma += valor;
                        }
                    }

                    int progreso = 40 + (int)((fila - 1f) / (ultimaFila - 1f) * 50); // De 40 a 90
                    reportar($"➕ Sumando BRUTO en fila {fila}...", progreso);
                }

                return $"💲 Suma total de BRUTO: ${suma:N2}";
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }


        public class MiniFurRow
        {
            public string RazonSocial { get; set; }
            public string NombreComercio { get; set; }
            public string CBU_CVU { get; set; }
            public string NroCuenta { get; set; }
            public string Legajo { get; set; }
            public string CUIT { get; set; }
            public string CodigoPostal { get; set; }
            public double Importe { get; set; }
        }
        public List<MiniFurRow> ObtenerFaltantesEnFUR(Action<string, int> reportar)
        {
            var resultadoDict = new Dictionary<string, MiniFurRow>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                reportar("🔎 Buscando grupos faltantes...", 0);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hojaFUR = workbook.Sheets["FUR"] as Excel.Worksheet;
                Excel.Worksheet hojaDiario = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;

                int filaDatosFUR = 10;
                int ultimaFilaFUR = hojaFUR.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // Mapeo columnas FUR
                int colRazon = 2, colComercio = 3, colCBU = 4, colCuenta = 5, colLegajo = 6, colCUIT = 7, colCP = 8;
                // Mapeo Diario2
                int colPagoReal = 57, colRazonD = 52, colComercioD = 44, colCBUD = 39, colCuentaD = 42, colLegajoD = 53, colCUITD = 32, colCPD = 54, colTotalDesc = 38;
                int ultimaFilaDiario = hojaDiario.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // 1. Guardar claves del FUR
                var clavesFUR = new HashSet<string>();
                for (int filaFUR = filaDatosFUR; filaFUR <= ultimaFilaFUR; filaFUR++)
                {
                    string razon = Convert.ToString((hojaFUR.Cells[filaFUR, colRazon] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string comercio = Convert.ToString((hojaFUR.Cells[filaFUR, colComercio] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string cbu = Convert.ToString((hojaFUR.Cells[filaFUR, colCBU] as Excel.Range)?.Text ?? "").Trim();
                    string cuenta = Convert.ToString((hojaFUR.Cells[filaFUR, colCuenta] as Excel.Range)?.Text ?? "").Trim();
                    string legajo = Convert.ToString((hojaFUR.Cells[filaFUR, colLegajo] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string cuit = Convert.ToString((hojaFUR.Cells[filaFUR, colCUIT] as Excel.Range)?.Text ?? "").Trim();
                    string cp = Convert.ToString((hojaFUR.Cells[filaFUR, colCP] as Excel.Range)?.Text ?? "").Trim();
                    string key = $"{razon}|{comercio}|{cbu}|{cuenta}|{legajo}|{cuit}|{cp}";
                    clavesFUR.Add(key);
                }

                // 2. Buscar y agrupar los faltantes en Diario2 (Pago Real = SI PAGA)
                for (int fila = 2; fila <= ultimaFilaDiario; fila++)
                {
                    string pagoReal = Convert.ToString((hojaDiario.Cells[fila, colPagoReal] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    if (pagoReal != "SI PAGA") continue;

                    string razon = Convert.ToString((hojaDiario.Cells[fila, colRazonD] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string comercio = Convert.ToString((hojaDiario.Cells[fila, colComercioD] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string cbu = Convert.ToString((hojaDiario.Cells[fila, colCBUD] as Excel.Range)?.Text ?? "").Trim();
                    string cuenta = Convert.ToString((hojaDiario.Cells[fila, colCuentaD] as Excel.Range)?.Text ?? "").Trim();
                    string legajo = Convert.ToString((hojaDiario.Cells[fila, colLegajoD] as Excel.Range)?.Text ?? "").Trim().ToUpper();
                    string cuit = Convert.ToString((hojaDiario.Cells[fila, colCUITD] as Excel.Range)?.Text ?? "").Trim();
                    string cp = Convert.ToString((hojaDiario.Cells[fila, colCPD] as Excel.Range)?.Text ?? "").Trim();

                    string key = $"{razon}|{comercio}|{cbu}|{cuenta}|{legajo}|{cuit}|{cp}";

                    if (!clavesFUR.Contains(key))
                    {
                        var celdaDesc = hojaDiario.Cells[fila, colTotalDesc] as Excel.Range;
                        string valorCelda = celdaDesc != null && celdaDesc.Value2 != null
                            ? Convert.ToString(celdaDesc.Value2).Replace("$", "").Replace(".", "").Replace(",", ".").Trim()
                            : "0";
                        double totalDesc = 0;
                        double.TryParse(valorCelda, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out totalDesc);

                        // Agrupa sumando el importe si ya existe
                        if (!resultadoDict.ContainsKey(key))
                        {
                            resultadoDict[key] = new MiniFurRow
                            {
                                RazonSocial = razon,
                                NombreComercio = comercio,
                                CBU_CVU = cbu,
                                NroCuenta = cuenta,
                                Legajo = legajo,
                                CUIT = cuit,
                                CodigoPostal = cp,
                                Importe = 0
                            };
                        }
                        resultadoDict[key].Importe += totalDesc;
                    }
                }

                // Devolver agrupado
                return new List<MiniFurRow>(resultadoDict.Values);
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        public void ExportarFaltantesAExcel(string rutaGuardar, List<MiniFurRow> filas)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.DisplayAlerts = false;
            Excel.Workbook wb = excelApp.Workbooks.Add();
            Excel.Worksheet ws = wb.Sheets[1];

            // Agrupar antes de exportar
            var filasAgrupadas = filas
                .GroupBy(f => new { f.RazonSocial, f.NombreComercio, f.CBU_CVU, f.NroCuenta, f.Legajo, f.CUIT, f.CodigoPostal })
                .Select(g => new MiniFurRow
                {
                    RazonSocial = g.Key.RazonSocial,
                    NombreComercio = g.Key.NombreComercio,
                    CBU_CVU = g.Key.CBU_CVU,
                    NroCuenta = g.Key.NroCuenta,
                    Legajo = g.Key.Legajo,
                    CUIT = g.Key.CUIT,
                    CodigoPostal = g.Key.CodigoPostal,
                    Importe = g.Sum(x => x.Importe)
                })
                .ToList();

            // Escribir encabezados
            ws.Cells[1, 1] = "Razon Social";
            ws.Cells[1, 2] = "Nombre Comercio";
            ws.Cells[1, 3] = "CBU/CVU";
            ws.Cells[1, 4] = "Nro de cuenta";
            ws.Cells[1, 5] = "Legajo";
            ws.Cells[1, 6] = "CUIT";
            ws.Cells[1, 7] = "Codigo Postal";
            ws.Cells[1, 8] = "Importe";

            int filaActual = 2;
            foreach (var f in filasAgrupadas)
            {
                ws.Cells[filaActual, 1] = f.RazonSocial;
                ws.Cells[filaActual, 2] = f.NombreComercio;
                ws.Cells[filaActual, 3] = f.CBU_CVU;
                ws.Cells[filaActual, 4].NumberFormat = "@"; // Forzar texto
                ws.Cells[filaActual, 4] = f.NroCuenta?.ToString();
                ws.Cells[filaActual, 5] = f.Legajo;
                ws.Cells[filaActual, 6] = f.CUIT;
                ws.Cells[filaActual, 7] = f.CodigoPostal;
                ws.Cells[filaActual, 8] = f.Importe;
                filaActual++;
            }

            wb.SaveAs(rutaGuardar);
            wb.Close();
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
        public string ControlarCostoTransaccional(Action<string, int> reportar)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            int COL_BRUTO = 8;    // H
            int COL_COSTO = 31;   // AE

            List<int> filasMal = new List<int>();

            try
            {
                reportar("💸 Validando costo transaccional...", 90);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hoja = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;

                if (hoja == null)
                    throw new Exception("No se encontró la hoja 'Reporte Diario2'.");

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    // ===== BRUTO =====
                    string brutoTxt = (hoja.Cells[fila, COL_BRUTO] as Excel.Range)?.Value2?.ToString();
                    double bruto = 0;

                    if (!string.IsNullOrWhiteSpace(brutoTxt))
                    {
                        brutoTxt = brutoTxt
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Replace("(", "-")
                            .Replace(")", "")
                            .Replace(".", "")
                            .Replace(",", ".")
                            .Trim();

                        double.TryParse(
                            brutoTxt,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out bruto
                        );
                    }

                    if (bruto <= 0)
                        continue;

                    // ===== COSTO TRANSACCIONAL (1.2%) =====
                    double costoEsperado = Math.Round(bruto * 0.012, 2);

                    string costoTxt = (hoja.Cells[fila, COL_COSTO] as Excel.Range)?.Value2?.ToString();
                    double costoHoja = 0;

                    if (!string.IsNullOrWhiteSpace(costoTxt))
                    {
                        costoTxt = costoTxt
                            .Replace("$", "")
                            .Replace(" ", "")
                            .Replace(".", "")
                            .Replace(",", ".")
                            .Trim();

                        double.TryParse(
                            costoTxt,
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out costoHoja
                        );
                    }

                    if (Math.Abs(costoHoja - costoEsperado) > 0.01)
                        filasMal.Add(fila);

                    int progreso = 90 + (int)((fila - 1f) / (ultimaFila - 1f) * 10);
                    reportar($"💸 Validando costo transaccional fila {fila}...", progreso);
                }

                return filasMal.Count == 0
                    ? "✅ Todos los costos transaccionales (AE) están correctos."
                    : $"❌ Error en costo transaccional en filas: {string.Join(", ", filasMal)}";
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }


        public string ValidarIIBB(Action<string, int> reportar, Dictionary<string, double> alicuotas)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            List<string> errores = new List<string>();

            // COLUMNAS (revisá si estos índices son correctos para tu Excel!)
            int COL_PROVINCIA = 34; // AH
            int COL_BASE = 29;      // AC
            int COL_IIBB = 35;      // AI

            try
            {
                reportar("🔍 Validando IIBB por provincia...", 0);
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;
                workbook = excelApp.Workbooks.Open(rutaArchivo);
                Excel.Worksheet hoja = workbook.Sheets["Reporte Diario2"] as Excel.Worksheet;
                if (hoja == null)
                    throw new Exception("No se encontró la hoja 'Reporte Diario2'.");

                int ultimaFila = hoja.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                for (int fila = 2; fila <= ultimaFila; fila++)
                {
                    string provincia = Convert.ToString((hoja.Cells[fila, COL_PROVINCIA] as Excel.Range)?.Text ?? "").Trim();
                    if (string.IsNullOrWhiteSpace(provincia)) continue;

                    // Buscar alícuota
                    if (!alicuotas.TryGetValue(provincia, out double porcentaje))
                    {
                        errores.Add($"❌ Fila {fila}: Provincia '{provincia}' no encontrada en tabla de alícuotas.");
                        continue;
                    }

                    // Leer valor base
                    var celdaBase = hoja.Cells[fila, COL_BASE] as Excel.Range;
                    double baseImponible = 0;
                    double.TryParse(
                        celdaBase?.Value2?.ToString()?.Replace("$", "").Replace(".", "").Replace(",", ".").Trim(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out baseImponible);

                    // Leer valor IIBB
                    var celdaIIBB = hoja.Cells[fila, COL_IIBB] as Excel.Range;
                    string valorIIBBStr = celdaIIBB?.Text?.Replace("$", "").Trim();

                    // Si alícuota es 0%
                    if (Math.Abs(porcentaje) < 0.0001)
                    {
                        if (!string.IsNullOrWhiteSpace(valorIIBBStr) && valorIIBBStr != "-")
                        {
                            errores.Add($"❌ Fila {fila}: Provincia '{provincia}' tiene alícuota 0%, pero en AI figura '{valorIIBBStr}' (debería ser '-' o vacío).");
                        }
                    }
                    else
                    {
                        // Calcular IIBB esperado
                        double esperado = Math.Round(baseImponible * porcentaje / 100, 2);

                        // Leer valor real
                        double valorIIBB = 0;
                        double.TryParse(
                            valorIIBBStr?.Replace(".", "").Replace(",", "."),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            out valorIIBB);

                        if (Math.Abs(valorIIBB - esperado) > 0.02) // tolera hasta 2 centavos
                        {
                            errores.Add($"❌ Fila {fila}: Provincia '{provincia}' IIBB calculado: {valorIIBB:N2}, esperado: {esperado:N2} ({porcentaje}%) sobre {baseImponible:N2}.");
                        }
                    }

                    // Reportar progreso cada 50 filas
                    if (fila % 50 == 0)
                        reportar($"🔍 Validando IIBB fila {fila}...", 50 + (int)(50.0 * fila / ultimaFila));
                }

                if (errores.Count == 0)
                    return "✅ Todos los IIBB están correctamente calculados.";
                else
                    return "❌ Errores en IIBB:\n" + string.Join("\n", errores);
            }
            finally
            {
                workbook?.Close(false);
                excelApp?.Quit();
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }
        private double ParseDouble(object value)
{
    if (value == null)
        return 0;

    string txt = Convert.ToString(value)
                 ?.Replace("$", "")
                 ?.Replace(" ", "")
                 ?.Replace("(", "-")
                 ?.Replace(")", "")
                 ?.Replace(".", "")
                 ?.Replace(",", ".")
                 ?.Trim();

    double.TryParse(
        txt,
        System.Globalization.NumberStyles.Any,
        System.Globalization.CultureInfo.InvariantCulture,
        out double result);

    return result;
}

    }
}