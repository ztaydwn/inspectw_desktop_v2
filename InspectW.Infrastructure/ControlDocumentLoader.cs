using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace InspectW.Infrastructure
{
    public class ControlDocumentLoader : IControlDocumentLoader
    {
        public Dictionary<int, string>? Load(IDictionary<string, byte[]> archivos)
        {
            if (archivos == null || archivos.Count == 0)
                return null;

            var candidate = archivos.Keys
                .FirstOrDefault(k => Path.GetFileName(k).StartsWith("control_documents", StringComparison.OrdinalIgnoreCase));

            if (candidate == null)
                return null;

            var data = archivos[candidate];
            var ext = Path.GetExtension(candidate).ToLowerInvariant();

            return ext switch
            {
                ".json" => ParseJson(data),
                ".csv" => ParseCsv(data),
                ".txt" => ParseTxt(data),
                _ => ParseTxt(data)
            };
        }

        private static Dictionary<int, string>? ParseJson(byte[] data)
        {
            try
            {
                var json = JsonDocument.Parse(data);
                var root = json.RootElement;
                var result = new Dictionary<int, string>();

                if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in root.EnumerateObject())
                    {
                        if (int.TryParse(prop.Name, out var num))
                            result[num] = prop.Value.ToString();
                    }
                    return result.Count > 0 ? result : null;
                }

                if (root.ValueKind == JsonValueKind.Array)
                {
                    foreach (var el in root.EnumerateArray())
                    {
                        int num;
                        if (el.TryGetProperty("numero", out var numProp) && numProp.TryGetInt32(out num))
                        {
                            var sit = el.TryGetProperty("situacion", out var sProp) ? sProp.ToString() : string.Empty;
                            result[num] = sit;
                        }
                        else if (el.TryGetProperty("num", out var numProp2) && numProp2.TryGetInt32(out num))
                        {
                            var sit = el.TryGetProperty("situacion", out var sProp) ? sProp.ToString() : string.Empty;
                            result[num] = sit;
                        }
                    }
                    return result.Count > 0 ? result : null;
                }
            }
            catch
            {
                return null;
            }

            return null;
        }

        private static Dictionary<int, string>? ParseCsv(byte[] data)
        {
            try
            {
                var text = Encoding.UTF8.GetString(data);
                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return null;

                var delimiter = DetectDelimiter(lines[0]);
                var headers = lines[0].Split(delimiter);
                var hasHeader = headers.Any(h => h.Equals("numero", StringComparison.OrdinalIgnoreCase) || h.StartsWith("num", StringComparison.OrdinalIgnoreCase));

                var result = new Dictionary<int, string>();
                for (var i = hasHeader ? 1 : 0; i < lines.Length; i++)
                {
                    var cols = lines[i].Split(delimiter);
                    if (cols.Length == 0) continue;

                    int num;
                    string situacion = string.Empty;
                    if (hasHeader)
                    {
                        var idxNum = Array.FindIndex(headers, h => h.StartsWith("num", StringComparison.OrdinalIgnoreCase));
                        if (idxNum >= 0 && idxNum < cols.Length && int.TryParse(cols[idxNum], out num))
                        {
                            var idxSit = Array.FindIndex(headers, h => h.StartsWith("sit", StringComparison.OrdinalIgnoreCase));
                            situacion = idxSit >= 0 && idxSit < cols.Length ? cols[idxSit] : cols.Last();
                            result[num] = situacion;
                        }
                    }
                    else if (int.TryParse(cols[0], out num))
                    {
                        situacion = cols.Length > 1 ? cols[1] : string.Empty;
                        result[num] = situacion;
                    }
                }
                return result.Count > 0 ? result : null;
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<int, string>? ParseTxt(byte[] data)
        {
            try
            {
                var text = Encoding.UTF8.GetString(data);
                var result = new Dictionary<int, string>();
                var pattern = new Regex(@"^\s*(\d{1,3})\b.*?SITUACION\s*:\s*(.+)$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
                foreach (Match m in pattern.Matches(text))
                {
                    if (int.TryParse(m.Groups[1].Value, out var num))
                        result[num] = m.Groups[2].Value.Trim();
                }

                if (result.Count == 0)
                {
                    var simple = new Regex(@"^\s*(\d{1,3})\s*[:.-]\s*(.+)$", RegexOptions.Multiline);
                    foreach (Match m in simple.Matches(text))
                    {
                        if (int.TryParse(m.Groups[1].Value, out var num))
                            result[num] = m.Groups[2].Value.Trim();
                    }
                }

                return result.Count > 0 ? result : null;
            }
            catch
            {
                return null;
            }
        }

        private static char DetectDelimiter(string sample)
        {
            if (sample.Contains(';')) return ';';
            if (sample.Contains('\t')) return '\t';
            return ',';
        }
    }
}
