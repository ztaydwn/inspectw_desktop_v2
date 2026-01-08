using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    public interface IGroupingService
    {
        Task<(IEnumerable<Grupo> grupos, List<string> errores)> AgruparAsync(IDictionary<string, byte[]> archivos, CancellationToken ct = default);
    }
}
