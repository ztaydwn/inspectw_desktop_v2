using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using InspectW.Domain;

namespace InspectW.Reporting
{
    public interface IXlsxReportService
    {
        Task GenerateAsync(
            IEnumerable<Grupo> grupos,
            IDictionary<string, byte[]> archivos,
            string destino,
            byte[]? plantilla = null,
            IDictionary<int, string>? controlDocs = null,
            IProgress<int>? progreso = null,
            CancellationToken ct = default);
    }
}
