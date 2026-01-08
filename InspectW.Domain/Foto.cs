namespace InspectW.Domain
{
    public class Foto
    {
        public string Archivo { get; set; } = string.Empty; // nombre de archivo
        public string Carpeta { get; set; } = string.Empty; // subcarpeta/origen
        public string Descripcion { get; set; } = string.Empty; // texto completo
        public string Detalle { get; set; } = string.Empty; // detalle extraído
        public byte[]? Bytes { get; set; } // opcional para thumbnails/reportes
    }
}
