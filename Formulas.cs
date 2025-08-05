using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel
{
    public class Formulas
    {
        public static void InsertarFormula(string tipoFormula, DataGridView dgv)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== INSERTAR FORMULA: {tipoFormula} ===");

                string rangoSeleccionado = new GestorSeleccion(dgv).ObtenerRangoSeleccionado();

                if (string.IsNullOrEmpty(rangoSeleccionado) || rangoSeleccionado.Trim() == "")
                {
                    rangoSeleccionado = $"{ConvertirACeldaRef(dgv.CurrentCell.ColumnIndex, dgv.CurrentCell.RowIndex)}";
                    System.Diagnostics.Debug.WriteLine($"Sin selección válida, usando celda actual: {rangoSeleccionado}");
                }
                else
                {
                    rangoSeleccionado = rangoSeleccionado.Trim();
                    System.Diagnostics.Debug.WriteLine($"Rango seleccionado limpio: '{rangoSeleccionado}'");
                }

                System.Diagnostics.Debug.WriteLine($"Rango final a usar: {rangoSeleccionado}");

                string formulaCompleta = $"={tipoFormula}({rangoSeleccionado})";
                System.Diagnostics.Debug.WriteLine($"Fórmula completa: {formulaCompleta}");

                var celdaDestino = EncontrarCeldaDestino(dgv, rangoSeleccionado);

                double resultado = Evaluar(formulaCompleta, dgv);
                System.Diagnostics.Debug.WriteLine($"Resultado de evaluación: {resultado}");

                if (!double.IsNaN(resultado) && !double.IsInfinity(resultado))
                {
                    celdaDestino.Value = resultado;
                    celdaDestino.Tag = formulaCompleta;
                    celdaDestino.Style.ForeColor = Color.Black;
                }
                else
                {
                    ManejarErrorFormula(celdaDestino, formulaCompleta, tipoFormula);
                }

                dgv.CurrentCell = celdaDestino;
                dgv.InvalidateCell(celdaDestino);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR en InsertarFormula: {ex.Message}");
                MessageBox.Show($"Error al insertar fórmula: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private static void ManejarErrorFormula(DataGridViewCell celdaDestino, string formulaCompleta, string tipoFormula)
        {
            if (EsOperacionDivision(tipoFormula))
            {
                MessageBox.Show("No es válido dividir por cero.", "Error de División",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning);
                celdaDestino.Value = "#DIV/0!";
                System.Diagnostics.Debug.WriteLine("Error de división por cero mostrado");
            }
            else
            {
                celdaDestino.Value = "#ERROR";
                System.Diagnostics.Debug.WriteLine("Error genérico mostrado");
            }

            celdaDestino.Style.ForeColor = Color.Red;
            celdaDestino.Tag = formulaCompleta;
        }
        private static bool EsOperacionDivision(string tipoFormula)
        {
            string formulaUpper = tipoFormula.ToUpper();
            return formulaUpper.Contains("DIVIDIR") ||
                   formulaUpper.Contains("DIVISION") ||
                   formulaUpper.Contains("DIVIDE");
        }
        private static DataGridViewCell EncontrarCeldaDestino(DataGridView dgv, string rango)
        {
            try
            {
                if (rango.Contains(":"))
                {
                    var partes = rango.Split(':');
                    string celdaFin = partes[1].Trim();

                    if (ParseCelda(celdaFin, out int col, out int fila))
                    {
                        fila--;
                        int filaDestino = fila + 1;
                        while (filaDestino < dgv.Rows.Count &&
                               dgv[col, filaDestino].Value != null &&
                               !string.IsNullOrWhiteSpace(dgv[col, filaDestino].Value.ToString()))
                        {
                            filaDestino++;
                        }

                        if (filaDestino >= dgv.Rows.Count)
                        {
                            filaDestino = dgv.Rows.Count - 1;
                        }

                        return dgv[col, filaDestino];
                    }
                }
                else
                {
                    if (ParseCelda(rango, out int col, out int fila))
                    {
                        fila--;
                        int filaDestino = Math.Min(fila + 1, dgv.Rows.Count - 1);
                        return dgv[col, filaDestino];
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EncontrarCeldaDestino: {ex.Message}");
            }

            return dgv.CurrentCell;
        }
        private static string ConvertirACeldaRef(int columna, int fila)
        {
            return $"{(char)('A' + columna)}{fila + 1}";
        }
        public static double Evaluar(string formula, DataGridView dgv)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(formula))
                {
                    System.Diagnostics.Debug.WriteLine("Fórmula vacía, retornando 0");
                    return 0;
                }

                if (!formula.StartsWith("="))
                {
                    System.Diagnostics.Debug.WriteLine("No es fórmula, convirtiendo a número");
                    return ConvertirTextoANumero(formula);
                }

                formula = formula.Substring(1).Trim().ToUpper();
                System.Diagnostics.Debug.WriteLine($"Fórmula procesada: '{formula}'");

                if (formula.StartsWith("SUMA(") || formula.StartsWith("SUM("))
                    return EvaluarSuma(formula, dgv);
                if (formula.StartsWith("RESTA(") || formula.StartsWith("SUBTRACT("))
                    return EvaluarResta(formula, dgv);
                if (formula.StartsWith("PRODUCTO(") || formula.StartsWith("MULTIPLICAR("))
                    return EvaluarProducto(formula, dgv);
                if (formula.StartsWith("DIVIDIR(") || formula.StartsWith("DIVISION("))
                    return EvaluarDivision(formula, dgv);

                if (formula.StartsWith("PROMEDIO(") || formula.StartsWith("AVERAGE("))
                    return EvaluarPromedio(formula, dgv);
                if (formula.StartsWith("MAX("))
                    return EvaluarMax(formula, dgv);
                if (formula.StartsWith("MIN("))
                    return EvaluarMin(formula, dgv);
                if (formula.StartsWith("COUNT("))
                    return EvaluarCount(formula, dgv);

                formula = ReemplazarReferencias(formula, dgv);
                return EvaluarExpresion(formula);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en Evaluar: {ex.Message}");
                return double.NaN;
            }
        }
        private static double EvaluarDivision(string formula, DataGridView dgv)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== INICIO EvaluarDivision: '{formula}' ===");

                string rango = ExtraerRango(formula);
                System.Diagnostics.Debug.WriteLine($"Rango extraído: '{rango}'");

                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("DIVIDIR: Rango vacío o inválido");
                    return double.NaN;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                System.Diagnostics.Debug.WriteLine($"DIVIDIR: Valores válidos: [{string.Join(", ", valoresValidos)}]");

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("DIVIDIR: No hay valores válidos");
                    return double.NaN;
                }

                if (valoresValidos.Count == 1)
                {
                    double valorUnico = valoresValidos[0];
                    System.Diagnostics.Debug.WriteLine($"DIVIDIR: Solo un valor: {valorUnico}");
                    return valorUnico;
                }

                double resultado = valoresValidos[0];
                System.Diagnostics.Debug.WriteLine($"DIVIDIR: Valor inicial: {resultado}");

                for (int i = 1; i < valoresValidos.Count; i++)
                {
                    System.Diagnostics.Debug.WriteLine($"DIVIDIR: Dividiendo {resultado} / {valoresValidos[i]}");

                    if (valoresValidos[i] == 0.0)
                    {
                        System.Diagnostics.Debug.WriteLine("DIVIDIR: División por cero detectada");
                        return double.NaN;
                    }

                    resultado /= valoresValidos[i];
                    System.Diagnostics.Debug.WriteLine($"DIVIDIR: Resultado parcial: {resultado}");
                }

                System.Diagnostics.Debug.WriteLine($"DIVIDIR: Resultado final = {resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR en EvaluarDivision: {ex.Message}");
                return double.NaN;
            }
        }
        private static double EvaluarPromedio(string formula, DataGridView dgv)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== INICIO EvaluarPromedio: '{formula}' ===");

                string rango = ExtraerRango(formula);
                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("PROMEDIO: Rango vacío o inválido");
                    return 0;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                System.Diagnostics.Debug.WriteLine($"PROMEDIO: Valores válidos: [{string.Join(", ", valoresValidos)}]");

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("PROMEDIO: No hay valores válidos, retornando 0");
                    return 0;
                }

                double resultado = valoresValidos.Average();
                System.Diagnostics.Debug.WriteLine($"PROMEDIO: Resultado final = {resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarPromedio: {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarProducto(string formula, DataGridView dgv)
        {
            try
            {
                string rango = ExtraerRango(formula);
                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("PRODUCTO: Rango vacío o inválido");
                    return double.NaN;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("PRODUCTO: No hay valores numéricos válidos");
                    return 0;
                }

                if (valoresValidos.All(v => v == 0))
                {
                    System.Diagnostics.Debug.WriteLine("PRODUCTO: Todos los valores son cero, resultado = 0");
                    return 0;
                }

                double producto = 1;
                foreach (var valor in valoresValidos)
                {
                    producto *= valor;
                    if (double.IsInfinity(producto))
                    {
                        System.Diagnostics.Debug.WriteLine("PRODUCTO: Overflow detectado");
                        return double.PositiveInfinity;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"PRODUCTO: valores=[{string.Join(",", valoresValidos)}], resultado={producto}");
                return producto;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarProducto: {ex.Message}");
                return double.NaN;
            }
        }
        private static double EvaluarCount(string formula, DataGridView dgv)
        {
            try
            {
                var rango = ExtraerRango(formula);
                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                int contador = valoresValidos.Count;

                System.Diagnostics.Debug.WriteLine($"COUNT: rango={rango}, valores numéricos válidos={contador}");
                return contador;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarCount: {ex.Message}");
                return 0;
            }
        }
        private static double ConvertirTextoANumero(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return 0;

            texto = texto.Trim();

            if (double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out double resultado))
                return resultado;

            if (double.TryParse(texto, NumberStyles.Any, CultureInfo.CurrentCulture, out resultado))
                return resultado;

            return 0;
        }
        private static string ReemplazarReferencias(string formula, DataGridView dgv)
        {
            var regex = new Regex(@"[A-Z]+\d+");
            var matches = regex.Matches(formula).Cast<Match>().ToList();

            matches = matches.OrderByDescending(m => m.Index).ToList();

            System.Diagnostics.Debug.WriteLine($"Formula original: {formula}");
            System.Diagnostics.Debug.WriteLine($"Referencias encontradas: {matches.Count}");

            foreach (Match match in matches)
            {
                var refCelda = match.Value;
                var valor = ObtenerValorCelda(refCelda, dgv);
                var valorStr = valor.ToString(CultureInfo.InvariantCulture);

                System.Diagnostics.Debug.WriteLine($"Reemplazando {refCelda} por {valorStr} en posición {match.Index}");

                formula = formula.Remove(match.Index, match.Length).Insert(match.Index, valorStr);

                System.Diagnostics.Debug.WriteLine($"Formula después del reemplazo: {formula}");
            }

            System.Diagnostics.Debug.WriteLine($"Formula final: {formula}");
            return formula;
        }
        private static double ObtenerValorCelda(string referencia, DataGridView dgv)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(referencia) || referencia.Length < 2)
                {
                    System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: Referencia inválida '{referencia}'");
                    return 0;
                }

                if (!ParseCelda(referencia, out int col, out int fila))
                {
                    System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: No se pudo parsear '{referencia}'");
                    return 0;
                }

                fila--;

                if (col < 0 || col >= dgv.Columns.Count || fila < 0 || fila >= dgv.Rows.Count)
                {
                    System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: Celda {referencia} fuera de rango [Col:{col}, Fila:{fila}]");
                    return 0;
                }

                var celda = dgv[col, fila];

                if (celda.Value == null)
                {
                    System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: Celda {referencia} está vacía");
                    return 0;
                }

                string valorTexto = celda.Value.ToString();

                if (celda.Tag is string tag && tag.StartsWith("="))
                {
                    double resultado = ConvertirTextoANumero(valorTexto);
                    System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: {referencia} (fórmula) = {resultado}");
                    return resultado;
                }

                double valor = ConvertirTextoANumero(valorTexto);
                System.Diagnostics.Debug.WriteLine($"ObtenerValorCelda: {referencia} = '{valorTexto}' -> {valor}");
                return valor;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR en ObtenerValorCelda({referencia}): {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarSuma(string formula, DataGridView dgv)
        {
            try
            {
                var rango = ExtraerRango(formula);
                var valores = ObtenerValoresRango(rango, dgv);
                double resultado = valores.Sum();
                System.Diagnostics.Debug.WriteLine($"SUMA: rango={rango}, valores=[{string.Join(",", valores)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarSuma: {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarResta(string formula, DataGridView dgv)
        {
            try
            {
                string rango = ExtraerRango(formula);
                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("RESTA: Rango vacío o inválido");
                    return double.NaN;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("RESTA: No hay valores numéricos válidos");
                    return 0;
                }

                if (valoresValidos.Count == 1)
                {
                    double resultado = valoresValidos[0];
                    System.Diagnostics.Debug.WriteLine($"RESTA: Un solo valor, resultado={resultado}");
                    return resultado;
                }

                double resta = valoresValidos[0];
                for (int i = 1; i < valoresValidos.Count; i++)
                {
                    resta -= valoresValidos[i];
                }

                System.Diagnostics.Debug.WriteLine($"RESTA: valores=[{string.Join(",", valoresValidos)}], resultado={resta}");
                return resta;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarResta: {ex.Message}");
                return double.NaN;
            }
        }
        private static double EvaluarMax(string formula, DataGridView dgv)
        {
            try
            {
                string rango = ExtraerRango(formula);
                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("MAX: Rango vacío o inválido");
                    return double.NaN;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("MAX: No hay valores numéricos válidos");
                    return 0;
                }

                double resultado = valoresValidos.Max();
                System.Diagnostics.Debug.WriteLine($"MAX: valores=[{string.Join(",", valoresValidos)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarMax: {ex.Message}");
                return double.NaN;
            }
        }
        private static double EvaluarMin(string formula, DataGridView dgv)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== INICIO EvaluarMin: '{formula}' ===");

                string rango = ExtraerRango(formula);
                if (string.IsNullOrEmpty(rango))
                {
                    System.Diagnostics.Debug.WriteLine("MIN: Rango vacío o inválido");
                    return double.NaN;
                }

                var valoresValidos = ObtenerValoresValidosParaOperacion(rango, dgv);

                if (!valoresValidos.Any())
                {
                    System.Diagnostics.Debug.WriteLine("MIN: No hay valores válidos");
                    return 0;
                }

                System.Diagnostics.Debug.WriteLine($"MIN: Valores válidos: [{string.Join(", ", valoresValidos)}]");

                double resultado = valoresValidos.Min();
                System.Diagnostics.Debug.WriteLine($"MIN: Resultado final = {resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR en EvaluarMin: {ex.Message}");
                return double.NaN;
            }
        }
        private static List<double> ObtenerValoresValidosParaOperacion(string rango, DataGridView dgv)
        {
            var valoresValidos = new List<double>();

            try
            {
                if (string.IsNullOrWhiteSpace(rango))
                    return valoresValidos;

                System.Diagnostics.Debug.WriteLine($"Obteniendo valores válidos para rango: {rango}");

                if (rango.Contains(":"))
                {
                    var partes = rango.Split(':');
                    if (partes.Length != 2) return valoresValidos;

                    string celdaInicio = partes[0].Trim();
                    string celdaFin = partes[1].Trim();

                    if (!ParseCelda(celdaInicio, out int colIni, out int filaIni) ||
                        !ParseCelda(celdaFin, out int colFin, out int filaFin))
                    {
                        return valoresValidos;
                    }

                    filaIni--;
                    filaFin--;

                    int minFila = Math.Min(filaIni, filaFin);
                    int maxFila = Math.Max(filaIni, filaFin);
                    int minCol = Math.Min(colIni, colFin);
                    int maxCol = Math.Max(colIni, colFin);

                    for (int f = minFila; f <= maxFila; f++)
                    {
                        for (int c = minCol; c <= maxCol; c++)
                        {
                            if (c >= 0 && c < dgv.Columns.Count && f >= 0 && f < dgv.Rows.Count)
                            {
                                var celda = dgv[c, f];

                                if (celda.Value != null && !string.IsNullOrWhiteSpace(celda.Value.ToString()))
                                {
                                    string valorTexto = celda.Value.ToString();
                                    double valor = ConvertirTextoANumero(valorTexto);

                                    if (!double.IsNaN(valor) && !double.IsInfinity(valor))
                                    {
                                        valoresValidos.Add(valor);
                                        System.Diagnostics.Debug.WriteLine($"Valor válido [{ColumnToLetter(c)}{f + 1}] = {valorTexto} -> {valor}");
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    var referencias = rango.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (var celda in referencias)
                    {
                        string celdaTrim = celda.Trim();
                        double valor = ObtenerValorCelda(celdaTrim, dgv);

                        if (!double.IsNaN(valor) && !double.IsInfinity(valor))
                        {
                            valoresValidos.Add(valor);
                            System.Diagnostics.Debug.WriteLine($"Referencia válida {celdaTrim} -> {valor}");
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Total valores válidos: {valoresValidos.Count} -> [{string.Join(", ", valoresValidos)}]");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en ObtenerValoresValidosParaOperacion: {ex.Message}");
            }

            return valoresValidos;
        }
        private static List<double> ObtenerValoresRango(string rango, DataGridView dgv)
        {
            var valores = new List<double>();

            try
            {
                if (string.IsNullOrWhiteSpace(rango))
                    return valores;

                System.Diagnostics.Debug.WriteLine($"Procesando rango: {rango}");

                if (rango.Contains(":"))
                {
                    var partes = rango.Split(':');
                    if (partes.Length != 2) return valores;

                    string celdaInicio = partes[0].Trim();
                    string celdaFin = partes[1].Trim();

                    System.Diagnostics.Debug.WriteLine($"Celda inicio: {celdaInicio}, Celda fin: {celdaFin}");

                    if (!ParseCelda(celdaInicio, out int colIni, out int filaIni) ||
                        !ParseCelda(celdaFin, out int colFin, out int filaFin))
                    {
                        System.Diagnostics.Debug.WriteLine("Error al parsear celdas del rango");
                        return valores;
                    }

                    System.Diagnostics.Debug.WriteLine($"Parsed - ColIni: {colIni}, FilaIni: {filaIni}, ColFin: {colFin}, FilaFin: {filaFin}");

                    filaIni--;
                    filaFin--;

                    int minFila = Math.Min(filaIni, filaFin);
                    int maxFila = Math.Max(filaIni, filaFin);
                    int minCol = Math.Min(colIni, colFin);
                    int maxCol = Math.Max(colIni, colFin);

                    System.Diagnostics.Debug.WriteLine($"Rango procesado: Filas {minFila + 1}-{maxFila + 1}, Columnas {ColumnToLetter(minCol)}-{ColumnToLetter(maxCol)}");

                    for (int f = minFila; f <= maxFila; f++)
                    {
                        for (int c = minCol; c <= maxCol; c++)
                        {
                            if (c >= 0 && c < dgv.Columns.Count && f >= 0 && f < dgv.Rows.Count)
                            {
                                var celda = dgv[c, f];
                                string valorTexto = celda.Value?.ToString() ?? "";

                                if (string.IsNullOrWhiteSpace(valorTexto))
                                {
                                    valores.Add(0);
                                    System.Diagnostics.Debug.WriteLine($"Celda [{ColumnToLetter(c)}{f + 1}] = (vacía) -> 0");
                                }
                                else
                                {
                                    double valor = ConvertirTextoANumero(valorTexto);
                                    valores.Add(valor);
                                    System.Diagnostics.Debug.WriteLine($"Celda [{ColumnToLetter(c)}{f + 1}] = {valorTexto} -> {valor}");
                                }
                            }
                            else
                            {
                                valores.Add(0);
                                System.Diagnostics.Debug.WriteLine($"Celda fuera de rango [{c},{f}] -> 0");
                            }
                        }
                    }
                }
                else
                {
                    var referencias = rango.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    System.Diagnostics.Debug.WriteLine($"Referencias individuales: {referencias.Length} elementos");

                    foreach (var celda in referencias)
                    {
                        string celdaTrim = celda.Trim();
                        double valor = ObtenerValorCelda(celdaTrim, dgv);
                        valores.Add(valor);
                        System.Diagnostics.Debug.WriteLine($"Referencia individual {celdaTrim} -> {valor}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Total de valores obtenidos: {valores.Count} -> [{string.Join(", ", valores)}]");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en ObtenerValoresRango: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
            }

            return valores;
        }
        private static bool ParseCelda(string celda, out int col, out int fila)
        {
            col = -1;
            fila = -1;

            if (string.IsNullOrWhiteSpace(celda) || celda.Length < 2)
                return false;

            try
            {
                int i = 0;
                while (i < celda.Length && char.IsLetter(celda[i]))
                {
                    i++;
                }

                if (i == 0 || i >= celda.Length)
                    return false;

                string columnaStr = celda.Substring(0, i);
                string filaStr = celda.Substring(i);

                col = 0;
                for (int j = 0; j < columnaStr.Length; j++)
                {
                    col = col * 26 + (columnaStr[j] - 'A' + 1);
                }
                col--;

                if (!int.TryParse(filaStr, out fila) || fila <= 0)
                    return false;

                System.Diagnostics.Debug.WriteLine($"ParseCelda({celda}) -> col={col} ({columnaStr}), fila={fila}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en ParseCelda({celda}): {ex.Message}");
                return false;
            }
        }
        private static string ColumnToLetter(int col)
        {
            string result = "";
            while (col >= 0)
            {
                result = (char)('A' + (col % 26)) + result;
                col = col / 26 - 1;
            }
            return result;
        }
        private static string ExtraerRango(string formula)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"ExtraerRango entrada: '{formula}'");

                int inicio = formula.IndexOf('(');
                int fin = formula.LastIndexOf(')');

                System.Diagnostics.Debug.WriteLine($"Posición '(': {inicio}, Posición ')': {fin}");

                if (inicio == -1 || fin == -1 || fin <= inicio)
                {
                    System.Diagnostics.Debug.WriteLine("ERROR: Paréntesis inválidos");
                    return string.Empty;
                }

                string rango = formula.Substring(inicio + 1, fin - inicio - 1).Trim();
                System.Diagnostics.Debug.WriteLine($"ExtraerRango resultado: '{rango}'");
                return rango;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ERROR en ExtraerRango: {ex.Message}");
                return string.Empty;
            }
        }
        private static double EvaluarExpresion(string expresion)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"Evaluando expresión: '{expresion}'");

                if (string.IsNullOrWhiteSpace(expresion))
                {
                    System.Diagnostics.Debug.WriteLine("Expresión vacía, retornando 0");
                    return 0;
                }

                var dt = new DataTable();
                var resultado = Convert.ToDouble(dt.Compute(expresion, null));

                if (double.IsInfinity(resultado) || double.IsNaN(resultado))
                {
                    System.Diagnostics.Debug.WriteLine($"Resultado inválido: {resultado}");
                    return double.NaN;
                }

                System.Diagnostics.Debug.WriteLine($"Resultado de la expresión: {resultado}");
                return resultado;
            }
            catch (DivideByZeroException)
            {
                System.Diagnostics.Debug.WriteLine("División por cero capturada en expresión");
                return double.NaN;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarExpresion({expresion}): {ex.Message}");
                return double.NaN;
            }
        }
    }
}