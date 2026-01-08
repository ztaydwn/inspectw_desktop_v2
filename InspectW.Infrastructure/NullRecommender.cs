using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    /// <summary>
    /// Placeholder que no modifica recomendaciones. Útil hasta que se integre el motor real.
    /// </summary>
    public class NullRecommender : IRecommender
    {
        public Task ApplyAsync(IEnumerable<Grupo> grupos, IDictionary<string, byte[]> archivos, CancellationToken ct = default)
        {
            return Task.CompletedTask;
        }
    }
}
