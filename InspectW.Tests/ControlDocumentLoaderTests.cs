using System.Collections.Generic;
using FluentAssertions;
using InspectW.Infrastructure;
using Xunit;

namespace InspectW.Tests
{
    public class ControlDocumentLoaderTests
    {
        [Fact]
        public void Load_ShouldParseCsv()
        {
            var csv = "numero;situacion\n1;OK\n2;REVISAR";
            var archivos = new Dictionary<string, byte[]>
            {
                ["control_documents.csv"] = System.Text.Encoding.UTF8.GetBytes(csv)
            };

            var loader = new ControlDocumentLoader();
            var map = loader.Load(archivos);

            map.Should().NotBeNull();
            map!.Count.Should().Be(2);
            map[1].Should().Be("OK");
        }
    }
}
