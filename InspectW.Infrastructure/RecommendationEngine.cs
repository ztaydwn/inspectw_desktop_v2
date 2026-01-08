using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    /// <summary>
    /// Port simplificado del motor de recomendaciones en Python: encuentra el TAG más cercano y puntúa observaciones.
    /// </summary>
    public class RecommendationEngine : IRecommender
    {
        private readonly List<Row> _rows;

        private RecommendationEngine(List<Row> rows)
        {
            _rows = rows;
        }

        public static RecommendationEngine FromCsv(byte[] bytes)
        {
            var rows = ParseRows(bytes);
            return new RecommendationEngine(rows);
        }

        public Task ApplyAsync(IEnumerable<Grupo> grupos, IDictionary<string, byte[]> archivos, CancellationToken ct = default)
        {
            if (grupos == null) return Task.CompletedTask;
            var csvPath = archivos.Keys.FirstOrDefault(k => k.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) && k.Contains("historico", StringComparison.OrdinalIgnoreCase));
            if (csvPath == null) return Task.CompletedTask;

            var rows = ParseRows(archivos[csvPath]);
            if (rows.Count == 0) return Task.CompletedTask;
            _rows.Clear();
            _rows.AddRange(rows);

            foreach (var grupo in grupos)
            {
                ct.ThrowIfCancellationRequested();
                if (grupo.Recomendaciones.Count > 0) continue;

                var extra = string.Join(", ", grupo.Fotos.Select(f => $"{f.Carpeta} {f.Detalle}".Trim())).Trim();
                var recs = Suggest(grupo.Nombre, extra, topK: 1, minScore: 0.2);
                foreach (var rec in recs)
                {
                    grupo.Recomendaciones.Add(new Recomendacion { Texto = rec.rec, Fuente = "historico.csv" });
                }
            }
            return Task.CompletedTask;
        }

        private IEnumerable<(double score, string rec)> Suggest(string query, string extra, int topK, double minScore)
        {
            if (_rows.Count == 0) return Array.Empty<(double, string)>();
            var wordRe = new Regex(@"[a-zA-Z0-9áéíóúÁÉÍÓÚñÑüÜ]+", RegexOptions.Compiled);
            var qTokens = Tokens(query, wordRe).ToHashSet(StringComparer.OrdinalIgnoreCase);

            var bestTag = _rows
                .GroupBy(r => r.Tag)
                .Select(g => new { Tag = g.Key, Score = ScoreTokens(qTokens, g.First().TagTokens) })
                .OrderByDescending(x => x.Score)
                .FirstOrDefault();

            if (bestTag == null || bestTag.Score < minScore)
                return Array.Empty<(double, string)>();

            var obsTokens = Tokens($"{query} {extra}", wordRe).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var candidates = _rows.Where(r => string.Equals(r.Tag, bestTag.Tag, StringComparison.OrdinalIgnoreCase));
            var scored = candidates
                .Select(r => (score: 0.6 * ScoreTokens(obsTokens, r.ObsTokens) + 0.4 * bestTag.Score, rec: r.Rec))
                .Where(x => x.score >= minScore)
                .OrderByDescending(x => x.score)
                .Take(topK)
                .ToList();

            return scored;
        }

        private static double ScoreTokens(HashSet<string> a, IEnumerable<string> b)
        {
            var bSet = b.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var inter = a.Intersect(bSet, StringComparer.OrdinalIgnoreCase).Count();
            var union = a.Count + bSet.Count - inter;
            return union == 0 ? 0.0 : inter / (double)union;
        }

        private static IEnumerable<string> Tokens(string text, Regex wordRe)
        {
            return wordRe.Matches(text ?? string.Empty).Select(m => m.Value.ToLowerInvariant());
        }

        private static string Decode(byte[] bytes)
        {
            try { return Encoding.UTF8.GetString(bytes); }
            catch { }
            try { return Encoding.GetEncoding("ISO-8859-1").GetString(bytes); }
            catch { }
            return Encoding.Default.GetString(bytes);
        }

        private static char DetectDelimiter(string sample)
        {
            if (sample.Contains(';')) return ';';
            if (sample.Contains('\t')) return '\t';
            return ',';
        }

        private static List<Row> ParseRows(byte[] bytes)
        {
            var text = Decode(bytes);
            var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length == 0) return new List<Row>();

            var delimiter = DetectDelimiter(lines[0]);
            var headers = lines[0].Split(delimiter).Select(h => h.Trim().ToUpperInvariant()).ToArray();
            var idxTag = Array.FindIndex(headers, h => h == "TAG");
            var idxObs = Array.FindIndex(headers, h => h.StartsWith("OBS", StringComparison.OrdinalIgnoreCase));
            var idxRec = Array.FindIndex(headers, h => h.StartsWith("RECOMEND", StringComparison.OrdinalIgnoreCase));
            var idxSrc = Array.FindIndex(headers, h => h == "FUENTE");
            if (idxTag < 0 || idxRec < 0) return new List<Row>();

            var wordRe = new Regex(@"[a-zA-Z0-9áéíóúÁÉÍÓÚñÑüÜ]+", RegexOptions.Compiled);
            var rows = new List<Row>();
            for (int i = 1; i < lines.Length; i++)
            {
                var cols = lines[i].Split(delimiter);
                if (cols.Length <= Math.Max(idxTag, idxRec)) continue;
                var tag = cols[idxTag].Trim();
                var obs = idxObs >= 0 && idxObs < cols.Length ? cols[idxObs].Trim() : string.Empty;
                var rec = idxRec >= 0 && idxRec < cols.Length ? cols[idxRec].Trim() : string.Empty;
                var src = idxSrc >= 0 && idxSrc < cols.Length ? cols[idxSrc].Trim() : string.Empty;
                if (string.IsNullOrWhiteSpace(tag) || string.IsNullOrWhiteSpace(rec)) continue;
                rows.Add(new Row
                {
                    Tag = tag,
                    Obs = obs,
                    Rec = rec,
                    Src = src,
                    TagTokens = Tokens(tag, wordRe),
                    ObsTokens = Tokens(obs, wordRe)
                });
            }
            return rows;
        }

        private sealed class Row
        {
            public string Tag { get; set; } = string.Empty;
            public string Obs { get; set; } = string.Empty;
            public string Rec { get; set; } = string.Empty;
            public string Src { get; set; } = string.Empty;
            public IEnumerable<string> TagTokens { get; set; } = Array.Empty<string>();
            public IEnumerable<string> ObsTokens { get; set; } = Array.Empty<string>();
        }
    }
}
