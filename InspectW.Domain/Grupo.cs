using System.Collections.Generic;

namespace InspectW.Domain
{
    public class Grupo
    {
        public int Numero { get; set; }
        public string Nombre { get; set; } = string.Empty;
        public List<Foto> Fotos { get; set; } = new();
        public List<Recomendacion> Recomendaciones { get; set; } = new();
    }
}
