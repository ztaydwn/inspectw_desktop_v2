using System.Collections.Generic;

namespace InspectW.Domain
{
    public class DescripcionParseResult
    {
        public List<Grupo> Grupos { get; set; } = new();
        public List<string> Errores { get; set; } = new();
    }
}
