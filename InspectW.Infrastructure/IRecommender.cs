using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    public interface IRecommender
    {
        Task ApplyAsync(IEnumerable<Grupo> grupos, IDictionary<string, byte[]> archivos, CancellationToken ct = default);
    }
}
