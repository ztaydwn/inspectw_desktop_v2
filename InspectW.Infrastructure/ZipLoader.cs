using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace InspectW.Infrastructure
{
    public class ZipLoader : IZipLoader
    {
        public async Task<IDictionary<string, byte[]>> LoadAsync(string zipPath, CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(zipPath);
            var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            await Task.Run(() =>
            {
                using var stream = File.OpenRead(zipPath);
                using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
                foreach (var entry in zip.Entries)
                {
                    ct.ThrowIfCancellationRequested();
                    if (string.IsNullOrEmpty(entry.FullName) || entry.FullName.EndsWith("/"))
                        continue; // skip directories

                    using var entryStream = entry.Open();
                    using var ms = new MemoryStream();
                    entryStream.CopyTo(ms);
                    var key = entry.FullName.Replace('\\', '/').ToLowerInvariant();
                    if (!result.ContainsKey(key))
                    {
                        result[key] = ms.ToArray();
                    }
                }
            }, ct).ConfigureAwait(false);

            return result;
        }
    }
}
