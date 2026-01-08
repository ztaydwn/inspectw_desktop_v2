using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace InspectW.Infrastructure
{
    public interface IFolderLoader
    {
        Task<IDictionary<string, byte[]>> LoadAsync(string folderPath, CancellationToken ct = default);
    }
}
