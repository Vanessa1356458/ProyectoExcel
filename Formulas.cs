using System;
using System.Collections.Generic;
using System.Data;
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
            if (dgv?.CurrentCell != null)
            {
                string formula = "";
                switch (tipoFormula)
                {
                    case "SUMA":
                        formula = "=SUMA(";
                        break;
                    case "PROMEDIO":
                        formula = "=PROMEDIO(";
                        break;
                    case "MAX":
                        formula = "=MAX(";
                        break;
                    case "MIN":
                        formula = "=MIN(";
                        break;
                    case "COUNT":
                        formula = "=COUNT(";
                        break;
                }

                dgv.CurrentCell.Tag = "FORMULA:" + formula;
                dgv.CurrentCell.Value = formula;

                var form = dgv.FindForm() as Form1;
                if (form != null && form.TxtFormula != null)
                {
                    form.TxtFormula.Text = formula;
                    form.TxtFormula.Focus();
                    form.TxtFormula.SelectionStart = formula.Length;
                }
                dgv.BeginEdit(true);
            }
        }
        public static double Evaluar(string formula, DataGridView dgv)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(formula)) return 0;

                if (!formula.StartsWith("="))
                {
                    return ConvertirTextoANumero(formula);
                }

                formula = formula.Substring(1).Trim().ToUpper();

                if (formula.StartsWith("SUMA(") || formula.StartsWith("SUM("))
                    return EvaluarSuma(formula, dgv);
                if (formula.StartsWith("RESTA(") || formula.StartsWith("SUBTRACT("))
                    return EvaluarResta(formula, dgv);
                if (formula.StartsWith("PROMEDIO(") || formula.StartsWith("AVERAGE("))
                    return EvaluarPromedio(formula, dgv);
                if (formula.StartsWith("MAX("))
                    return EvaluarMax(formula, dgv);
                if (formula.StartsWith("MIN("))
                    return EvaluarMin(formula, dgv);

                formula = ReemplazarReferencias(formula, dgv);
                return EvaluarExpresion(formula);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en Evaluar: {ex.Message}");
                return double.NaN;
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
                    return 0;

                if (!ParseCelda(referencia, out int col, out int fila))
                    return 0;

                fila--; 

                if (col < 0 || col >= dgv.Columns.Count || fila < 0 || fila >= dgv.Rows.Count)
                    return 0;

                var celda = dgv[col, fila];
                if (celda.Value == null)
                    return 0;

                string valorTexto = celda.Value.ToString();

                if (celda.Tag is string tag && tag.StartsWith("="))
                {
                    return ConvertirTextoANumero(valorTexto);
                }

                return ConvertirTextoANumero(valorTexto);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en ObtenerValorCelda({referencia}): {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarSuma(string formula, DataGridView dgv)
        {
            try
            {
                var rango = ExtraerRango(formula);
                var valores = ObtenerValoresRango(rango, dgv);
                System.Diagnostics.Debug.WriteLine($"SUMA: rango={rango}, valores=[{string.Join(",", valores)}], resultado={valores.Sum()}");
                return valores.Sum();
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
                var rango = ExtraerRango(formula);
                var valores = ObtenerValoresRango(rango, dgv);

                if (!valores.Any())
                {
                    System.Diagnostics.Debug.WriteLine("RESTA: No hay valores para procesar");
                    return 0;
                }

                if (valores.Count == 1)
                {
                    System.Diagnostics.Debug.WriteLine($"RESTA: Solo un valor {valores[0]}");
                    return valores[0];
                }

                double resultado = valores[0];
                System.Diagnostics.Debug.WriteLine($"RESTA: Valor inicial = {resultado}");

                for (int i = 1; i < valores.Count; i++)
                {
                    System.Diagnostics.Debug.WriteLine($"RESTA: {resultado} - {valores[i]} = {resultado - valores[i]}");
                    resultado -= valores[i];
                }

                System.Diagnostics.Debug.WriteLine($"RESTA: rango={rango}, valores=[{string.Join(",", valores)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarResta: {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarPromedio(string formula, DataGridView dgv)
        {
            try
            {
                var valores = ObtenerValoresRango(ExtraerRango(formula), dgv);
                double resultado = valores.Any() ? valores.Average() : 0;
                System.Diagnostics.Debug.WriteLine($"PROMEDIO: valores=[{string.Join(",", valores)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarPromedio: {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarMax(string formula, DataGridView dgv)
        {
            try
            {
                var valores = ObtenerValoresRango(ExtraerRango(formula), dgv);
                double resultado = valores.Any() ? valores.Max() : 0;
                System.Diagnostics.Debug.WriteLine($"MAX: valores=[{string.Join(",", valores)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarMax: {ex.Message}");
                return 0;
            }
        }
        private static double EvaluarMin(string formula, DataGridView dgv)
        {
            try
            {
                var valores = ObtenerValoresRango(ExtraerRango(formula), dgv);
                double resultado = valores.Any() ? valores.Min() : 0;
                System.Diagnostics.Debug.WriteLine($"MIN: valores=[{string.Join(",", valores)}], resultado={resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarMin: {ex.Message}");
                return 0;
            }
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
            int inicio = formula.IndexOf('(') + 1;
            int fin = formula.LastIndexOf(')');
            string rango = (inicio >= 0 && fin > inicio) ? formula.Substring(inicio, fin - inicio) : "";
            System.Diagnostics.Debug.WriteLine($"ExtraerRango({formula}) -> {rango}");
            return rango;
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

                System.Diagnostics.Debug.WriteLine($"Resultado de la expresión: {resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error en EvaluarExpresion({expresion}): {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"StackTrace: {ex.StackTrace}");
                return double.NaN;
            }
        }
    }
}