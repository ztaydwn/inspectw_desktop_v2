using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace InspectW.Infrastructure
{
    public interface IZipLoader
    {
        Task<IDictionary<string, byte[]>> LoadAsync(string zipPath, CancellationToken ct = default);
    }
}
