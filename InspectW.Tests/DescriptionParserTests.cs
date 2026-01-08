using System.Collections.Generic;
using System.Threading.Tasks;
using FluentAssertions;
using InspectW.Infrastructure;
using Xunit;

namespace InspectW.Tests
{
    public class DescriptionParserTests
    {
        [Fact]
        public void Parse_ShouldBuildGroupsAndFotos()
        {
            var archivos = new Dictionary<string, byte[]>
            {
                ["descriptions.txt"] = @"[Paredes] foto1.jpg
Description: 1 Grieta leve
Recommendation: Revisar
".NormalizeLineEndings(),
                ["grupos.txt"] = "1 Paredes".NormalizeLineEndings(),
                ["Paredes/foto1.jpg".ToLower()] = new byte[] { 1, 2, 3 }
            };

            var parser = new DescriptionParser();
            var result = parser.Parse(archivos);

            result.Errores.Should().BeEmpty();
            result.Grupos.Should().HaveCount(1);
            result.Grupos[0].Fotos.Should().HaveCount(1);
            result.Grupos[0].Recomendaciones.Should().NotBeEmpty();
        }
    }

    internal static class TestStringExtensions
    {
        public static byte[] NormalizeLineEndings(this string s) => System.Text.Encoding.UTF8.GetBytes(s.Replace("\r\n", "\n"));
    }
}
