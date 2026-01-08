using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    public class GroupingService : IGroupingService
    {
        private readonly IDescriptionParser _parser;
        private readonly IRecommender? _recommender;

        public GroupingService(IDescriptionParser parser, IRecommender? recommender = null)
        {
            _parser = parser;
            _recommender = recommender;
        }

        public async Task<(IEnumerable<Grupo> grupos, List<string> errores)> AgruparAsync(IDictionary<string, byte[]> archivos, CancellationToken ct = default)
        {
            var errores = new List<string>();
            var parseResult = _parser.Parse(archivos);
            errores.AddRange(parseResult.Errores);

            foreach (var grupo in parseResult.Grupos)
            {
                foreach (var foto in grupo.Fotos)
                {
                    ct.ThrowIfCancellationRequested();
                    foto.Bytes = FindImageData(archivos, foto, errores);
                }
            }

            if (_recommender != null)
            {
                try
                {
                    await _recommender.ApplyAsync(parseResult.Grupos, archivos, ct).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    errores.Add($"Error al aplicar recomendaciones: {ex.Message}");
                }
            }

            return (parseResult.Grupos, errores);
        }

        private static byte[]? FindImageData(IDictionary<string, byte[]> archivos, Foto foto, List<string> errores)
        {
            var ideal = $"{foto.Carpeta}/{foto.Archivo}".Replace("\\", "/").Trim('/').ToLowerInvariant();
            if (archivos.TryGetValue(ideal, out var bytes))
                return bytes;

            // Buscar por nombre de archivo y, si es único, corregir carpeta
            var matches = archivos
                .Where(kv => string.Equals(Path.GetFileName(kv.Key), foto.Archivo, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (matches.Count == 1)
            {
                var path = matches[0].Key.Replace("\\", "/");
                foto.Carpeta = Path.GetDirectoryName(path)?.Replace("\\", "/")?.Trim('/') ?? string.Empty;
                return matches[0].Value;
            }

            if (matches.Count > 1)
            {
                var list = string.Join(", ", matches.Take(3).Select(m => m.Key));
                errores.Add($"Archivo '{foto.Archivo}' es ambiguo ({matches.Count} coincidencias: {list}...).");
            }
            else
            {
                errores.Add($"No se encontro la imagen '{foto.Archivo}'.");
            }

            return null;
        }
    }
}
