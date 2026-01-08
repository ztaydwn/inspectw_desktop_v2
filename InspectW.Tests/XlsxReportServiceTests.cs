using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using InspectW.Domain;
using InspectW.Reporting;
using Xunit;

namespace InspectW.Tests
{
    public class XlsxReportServiceTests
    {
        [Fact]
        public async Task GenerateAsync_ShouldCreateFile()
        {
            var grupos = new List<Grupo>
            {
                new Grupo
                {
                    Numero = 1,
                    Nombre = "1 Paredes",
                    Fotos = new List<Foto>
                    {
                        new Foto { Archivo = "f1.jpg", Detalle = "Grieta", Carpeta = "Paredes" }
                    },
                    Recomendaciones = new List<Recomendacion> { new Recomendacion { Texto = "Sellar" } }
                }
            };

            var archivos = new Dictionary<string, byte[]>();
            var svc = new XlsxReportService();
            var tmp = Path.Combine(Path.GetTempPath(), "reporte_test.xlsx");

            if (File.Exists(tmp))
                File.Delete(tmp);

            await svc.GenerateAsync(grupos, archivos, tmp);

            File.Exists(tmp).Should().BeTrue();
        }
    }
}
