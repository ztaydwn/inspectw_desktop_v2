using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    public class DescriptionParser : IDescriptionParser
    {
        private static readonly Regex HeaderRegex = new(@"\[(.+?)\]\s+(\S+\.jpg)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public DescripcionParseResult Parse(IDictionary<string, byte[]> archivos)
        {
            var result = new DescripcionParseResult();

            if (archivos == null || archivos.Count == 0)
            {
                result.Errores.Add("No se recibieron archivos para parsear.");
                return result;
            }

            var descriptions = TryReadText(archivos, "descriptions.txt");
            var gruposTxt = TryReadText(archivos, "grupos.txt");

            if (string.IsNullOrWhiteSpace(descriptions) || string.IsNullOrWhiteSpace(gruposTxt))
            {
                result.Errores.Add("Faltan 'descriptions.txt' o 'grupos.txt' en el origen de datos.");
                return result;
            }

            var groupLookup = CreateGroupLookup(gruposTxt);
            var groups = new Dictionary<string, Grupo>(StringComparer.OrdinalIgnoreCase);

            var lines = descriptions.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var current = new BlockState();
            for (var i = 0; i < lines.Length; i++)
            {
                var lineNum = i + 1;
                var line = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var m = HeaderRegex.Match(line);
                if (m.Success)
                {
                    // flush previous block
                    TryFlushBlock(current, groupLookup, groups, result.Errores, lineNum);
                    current = new BlockState
                    {
                        Carpeta = m.Groups[1].Value,
                        FileName = m.Groups[2].Value
                    };
                    continue;
                }

                if (line.StartsWith("description:", StringComparison.OrdinalIgnoreCase))
                {
                    var descContent = line.Split(':', 2)[1].Trim();
                    var descParts = Regex.Split(descContent, @"\s+", RegexOptions.None, TimeSpan.FromSeconds(1));
                    current.NumberingCode = descParts.Length > 0 ? descParts[0] : string.Empty;
                    current.Detail = descParts.Length > 1 ? string.Join(' ', descParts.Skip(1)).Trim() : string.Empty;
                    continue;
                }

                if (line.StartsWith("recommendation:", StringComparison.OrdinalIgnoreCase))
                {
                    current.Recommendation = line.Split(':', 2)[1].Trim();
                }
            }

            // flush last block
            TryFlushBlock(current, groupLookup, groups, result.Errores, lines.Length);

            result.Grupos = groups.Values.ToList();
            return result;
        }

        private static void TryFlushBlock(BlockState block, Dictionary<string, GroupInfo> lookup, Dictionary<string, Grupo> groups, List<string> errores, int lineNum)
        {
            if (!block.IsValid())
                return;

            var official = lookup.TryGetValue(block.NumberingCode, out var info)
                ? info
                : new GroupInfo { Key = block.NumberingCode, Name = $"Grupo no encontrado para '{block.NumberingCode}'", Numero = 0 };

            var grupo = groups.TryGetValue(official.Name, out var g)
                ? g
                : groups[official.Name] = new Grupo { Nombre = official.Name, Numero = official.Numero };

            var foto = new Foto
            {
                Archivo = block.FileName,
                Carpeta = NormalizeFolder(block.Carpeta),
                Detalle = block.Detail,
                Descripcion = block.Detail, // mantener detalle como descripcion base
            };

            grupo.Fotos.Add(foto);

            if (!string.IsNullOrWhiteSpace(block.Recommendation))
            {
                grupo.Recomendaciones.Add(new Recomendacion { Texto = block.Recommendation });
            }

            // warnings si no se encontro grupo
            if (official.Numero == 0 && !lookup.ContainsKey(block.NumberingCode))
            {
                errores.Add($"Linea {lineNum}: Codigo '{block.NumberingCode}' no existe en grupos.txt.");
            }
        }

        private static Dictionary<string, GroupInfo> CreateGroupLookup(string gruposTxt)
        {
            var lookup = new Dictionary<string, GroupInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in gruposTxt.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                var full = line.Trim();
                if (string.IsNullOrEmpty(full))
                    continue;
                if (!char.IsDigit(full[0])) // omitir cabeceras tipo ENUMERACION GRUPOS
                    continue;

                var parts = new Regex(@"\s+").Split(full, 2);
                if (parts.Length < 2)
                    continue;

                var key = parts[0].Trim();
                var name = $"{key} {parts[1].Trim()}";
                var numero = int.TryParse(key, out var n) ? n : 0;
                lookup[key] = new GroupInfo { Key = key, Name = name, Numero = numero };
            }
            return lookup;
        }

        private static string TryReadText(IDictionary<string, byte[]> archivos, string key)
        {
            if (!archivos.TryGetValue(key, out var bytes))
                return string.Empty;

            try { return Encoding.UTF8.GetString(bytes); }
            catch
            {
                try { return Encoding.GetEncoding("ISO-8859-1").GetString(bytes); }
                catch { return string.Empty; }
            }
        }

        private static string NormalizeFolder(string carpeta)
        {
            var c = carpeta?.Replace('\\', '/') ?? string.Empty;
            return c.Trim('/');
        }

        private sealed class GroupInfo
        {
            public string Key { get; set; } = string.Empty;
            public string Name { get; set; } = string.Empty;
            public int Numero { get; set; }
        }

        private sealed class BlockState
        {
            public string Carpeta { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public string NumberingCode { get; set; } = string.Empty;
            public string Detail { get; set; } = string.Empty;
            public string Recommendation { get; set; } = string.Empty;

            public bool IsValid()
            {
                return !string.IsNullOrWhiteSpace(Carpeta) && !string.IsNullOrWhiteSpace(FileName);
            }
        }
    }
}
