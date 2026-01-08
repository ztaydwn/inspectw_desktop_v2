using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace InspectW.Infrastructure
{
    public class FolderLoader : IFolderLoader
    {
        public async Task<IDictionary<string, byte[]>> LoadAsync(string folderPath, CancellationToken ct = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(folderPath);
            var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            var files = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                ct.ThrowIfCancellationRequested();
                var rel = Path.GetRelativePath(folderPath, file).Replace('\\', '/').ToLowerInvariant();
                if (result.ContainsKey(rel))
                    continue;

                var bytes = await File.ReadAllBytesAsync(file, ct).ConfigureAwait(false);
                result[rel] = bytes;
            }

            return result;
        }
    }
}
