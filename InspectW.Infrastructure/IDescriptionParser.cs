using System.Collections.Generic;
using InspectW.Domain;

namespace InspectW.Infrastructure
{
    public interface IDescriptionParser
    {
        DescripcionParseResult Parse(IDictionary<string, byte[]> archivos);
    }
}
