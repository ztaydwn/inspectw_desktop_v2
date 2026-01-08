using System.Collections.Generic;

namespace InspectW.Infrastructure
{
    public interface IControlDocumentLoader
    {
        Dictionary<int, string>? Load(IDictionary<string, byte[]> archivos);
    }
}
