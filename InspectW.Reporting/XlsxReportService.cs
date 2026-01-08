using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using InspectW.Domain;

namespace InspectW.Reporting
{
    public class XlsxReportService : IXlsxReportService
    {
        private static readonly XLColor HeaderFill = XLColor.FromArgb(217, 217, 217);
        private static readonly XLBorderStyleValues Thin = XLBorderStyleValues.Thin;

        public async Task GenerateAsync(
            IEnumerable<Grupo> grupos,
            IDictionary<string, byte[]> archivos,
            string destino,
            byte[]? plantilla = null,
            IDictionary<int, string>? controlDocs = null,
            IProgress<int>? progreso = null,
            CancellationToken ct = default)
        {
            ArgumentNullException.ThrowIfNull(grupos);
            ArgumentException.ThrowIfNullOrWhiteSpace(destino);

            await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                using var wb = plantilla is not null ? new XLWorkbook(new MemoryStream(plantilla)) : new XLWorkbook();
                if (wb.Worksheets.Count == 1 && wb.Worksheet(1).Name.Equals("Sheet1", StringComparison.OrdinalIgnoreCase))
                {
                    wb.Worksheets.Delete(1);
                }

                // Portada / Datos generales / Desarrollo (vacías si no hay info)
                var info = ExtractInfo(archivos);
                AddPortada(wb, info, archivos);
                AddDatosGenerales(wb, info);
                AddDesarrollo(wb);

                // Hojas por grupo con grid de fotos + tabla
                AddGroupSheets(wb, grupos, archivos, controlDocs, progreso, ct);

                // Control documents
                if (controlDocs != null && controlDocs.Count > 0)
                {
                    AddControlDocsSheet(wb, controlDocs);
                }

                Directory.CreateDirectory(Path.GetDirectoryName(destino) ?? ".");
                wb.SaveAs(destino);
            }, ct).ConfigureAwait(false);
        }

        private static void AddPortada(XLWorkbook wb, Dictionary<string, string> info, IDictionary<string, byte[]> archivos)
        {
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("Portada", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheets.Add("Portada");
            ws.Cell("A1").Value = "Informe de Inspección";
            ws.Cell("A1").Style.Font.SetBold();
            ws.Cell("A1").Style.Font.FontSize = 18;
            ws.Cell("A3").Value = "Proyecto / Establecimiento:";
            ws.Cell("B3").Value = info.GetValueOrDefault("establecimiento", info.GetValueOrDefault("titulo", ""));
            ws.Cell("A4").Value = "Propietario:";
            ws.Cell("B4").Value = info.GetValueOrDefault("propietario", "");
            ws.Cell("A5").Value = "Dirección:";
            ws.Cell("B5").Value = info.GetValueOrDefault("direccion", "");
            ws.Cell("A6").Value = "Fecha:";
            ws.Cell("B6").Value = info.GetValueOrDefault("fecha", DateTime.Today.ToShortDateString());
            ws.Columns().AdjustToContents();

            // Logo opcional (portada.png / portadat.png)
            var logoBytes = FindLogoBytes(archivos);
            if (logoBytes != null)
            {
                using var ms = new MemoryStream(logoBytes);
                var pic = ws.AddPicture(ms);
                pic.MoveTo(ws.Cell("D1"));
                pic.Scale(0.35);
            }
        }

        private static void AddDatosGenerales(XLWorkbook wb, Dictionary<string, string> info)
        {
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("Datos Generales", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheets.Add("Datos Generales");
            ws.Cell("A1").Value = "Clave";
            ws.Cell("B1").Value = "Valor";
            ws.Row(1).Style.Font.SetBold();
            int row = 2;
            foreach (var kv in info)
            {
                ws.Cell(row, 1).Value = kv.Key;
                ws.Cell(row, 2).Value = kv.Value;
                row++;
            }
            ws.Columns().AdjustToContents();
        }

        private static void AddDesarrollo(XLWorkbook wb)
        {
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("Desarrollo", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheets.Add("Desarrollo");
            ws.Cell("A1").Value = "Sección";
            ws.Cell("B1").Value = "Descripción";
            ws.Row(1).Style.Font.SetBold();
        }

        private static void AddGroupSheets(
            XLWorkbook wb,
            IEnumerable<Grupo> grupos,
            IDictionary<string, byte[]> archivos,
            IDictionary<int, string>? controlDocs,
            IProgress<int>? progreso,
            CancellationToken ct)
        {
            var ordered = grupos
                .OrderBy(g => g.Numero == 0 ? int.MaxValue : g.Numero)
                .ThenBy(g => g.Nombre)
                .ToList();

            var totalFotos = ordered.Sum(g => g.Fotos.Count);
            int processed = 0;

            foreach (var grupo in ordered)
            {
                ct.ThrowIfCancellationRequested();
                var sheetName = SanitizeSheetName(grupo.Nombre);
                var ws = wb.Worksheets.Any(w => w.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    ? wb.AddWorksheet($"{sheetName}_{Guid.NewGuid():N}".Substring(0, 6))
                    : wb.AddWorksheet(sheetName);

                // Configuración básica de página
                ws.PageSetup.PageOrientation = XLPageOrientation.Portrait;
                ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                ws.PageSetup.Margins.Left = 0.25;
                ws.PageSetup.Margins.Right = 0.25;
                ws.PageSetup.Margins.Top = 0.5;
                ws.PageSetup.Margins.Bottom = 0.5;

                // Banner
                ws.Range("A1:C1").Merge();
                var banner = ws.Cell("A1");
                banner.Value = "CONDICIÓN DE SEGURIDAD OBSERVADA:\nSEGÚN TABLA DE D.S. 007-2018-PCM (ANEXO 7A)";
                banner.Style.Alignment.WrapText = true;
                banner.Style.Font.SetBold();
                banner.Style.Fill.BackgroundColor = HeaderFill;
                banner.Style.Border.OutsideBorder = Thin;
                ws.Row(1).Height = 36;

                // Título
                ws.Range("A2:C2").Merge();
                var title = ws.Cell("A2");
                title.Value = grupo.Nombre;
                title.Style.Font.SetBold();
                title.Style.Font.FontSize = 14;

                // Grid de fotos (3 columnas)
                int startRow = 4;
                int cols = 3;
                double targetWidth = 180; // px aproximado
                double targetHeight = 140;
                for (int c = 0; c < cols; c++)
                {
                    ws.Column(c + 1).Width = 28;
                }

                for (int i = 0; i < grupo.Fotos.Count; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var foto = grupo.Fotos[i];
                    var rowBase = startRow + (i / cols) * 5;
                    var col = (i % cols) + 1;

                    var bytes = foto.Bytes ?? FindBytes(archivos, foto);
                    if (bytes != null)
                    {
                        using var ms = new MemoryStream(bytes);
                        var pic = ws.AddPicture(ms);
                        pic.MoveTo(ws.Cell(rowBase, col));
                        var scaleH = targetHeight / pic.OriginalHeight;
                        var scaleW = targetWidth / pic.OriginalWidth;
                        var scale = Math.Min(scaleH, scaleW);
                        pic.Scale(scale);
                    }
                    ws.Cell(rowBase + 1, col).Value = $"{foto.Carpeta} - {foto.Detalle}";
                    ws.Row(rowBase + 1).Height = 32;
                }

                // Tabla de detalle/recomendación
                int tableStart = startRow + ((grupo.Fotos.Count + cols - 1) / cols) * 5 + 2;
                var headers = new List<string> { "Carpeta", "Archivo", "Detalle", "Recomendación" };
                var hasControl = controlDocs != null && controlDocs.Count > 0;
                if (hasControl)
                    headers.Insert(3, "Situación");

                for (int i = 0; i < headers.Count; i++)
                {
                    ws.Cell(tableStart, i + 1).Value = headers[i];
                    ws.Cell(tableStart, i + 1).Style.Font.Bold = true;
                    ws.Cell(tableStart, i + 1).Style.Fill.BackgroundColor = HeaderFill;
                    ws.Cell(tableStart, i + 1).Style.Border.OutsideBorder = Thin;
                }

                int row = tableStart + 1;
                foreach (var foto in grupo.Fotos)
                {
                    ct.ThrowIfCancellationRequested();
                    ws.Cell(row, 1).Value = foto.Carpeta;
                    ws.Cell(row, 2).Value = foto.Archivo;
                    ws.Cell(row, 3).Value = foto.Detalle;
                    if (hasControl)
                    {
                        var situacion = controlDocs!.TryGetValue(grupo.Numero, out var sit) ? sit : string.Empty;
                        ws.Cell(row, 4).Value = situacion;
                        ws.Cell(row, 5).Value = grupo.Recomendaciones.FirstOrDefault()?.Texto ?? string.Empty;
                    }
                    else
                    {
                        ws.Cell(row, 4).Value = grupo.Recomendaciones.FirstOrDefault()?.Texto ?? string.Empty;
                    }
                    // bordes fila
                    for (int c = 1; c <= headers.Count; c++)
                    {
                        ws.Cell(row, c).Style.Border.OutsideBorder = Thin;
                    }
                    row++;
                    processed++;
                    progreso?.Report(ComputePercent(processed, totalFotos));
                }

                ws.Columns().AdjustToContents();
            }
        }

        private static byte[]? FindBytes(IDictionary<string, byte[]> archivos, Foto foto)
        {
            var ideal = $"{foto.Carpeta}/{foto.Archivo}".Replace("\\", "/").Trim('/').ToLowerInvariant();
            if (archivos.TryGetValue(ideal, out var bytes))
                return bytes;
            var match = archivos.FirstOrDefault(kv => kv.Key.EndsWith("/" + foto.Archivo, StringComparison.OrdinalIgnoreCase));
            return match.Equals(default(KeyValuePair<string, byte[]>)) ? null : match.Value;
        }

        private static void AddControlDocsSheet(XLWorkbook wb, IDictionary<int, string> controlDocs)
        {
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("ControlDocs", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheets.Add("ControlDocs");
            ws.Cell("A1").Value = "Número";
            ws.Cell("B1").Value = "Situación";
            ws.Row(1).Style.Font.SetBold();
            int row = 2;
            foreach (var kv in controlDocs.OrderBy(k => k.Key))
            {
                ws.Cell(row, 1).Value = kv.Key;
                ws.Cell(row, 2).Value = kv.Value;
                row++;
            }
            ws.Columns().AdjustToContents();
        }

        private static string SanitizeSheetName(string name)
        {
            var invalid = new[] { "/", "\\", "?", "*", "[", "]", ":" };
            foreach (var ch in invalid)
            {
                name = name.Replace(ch, "-");
            }
            name = (name ?? "Hoja").Trim();
            if (string.IsNullOrEmpty(name)) name = "Hoja";
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }

        private static Dictionary<string, string> ExtractInfo(IDictionary<string, byte[]> archivos)
        {
            foreach (var kv in archivos)
            {
                var baseName = Path.GetFileName(kv.Key).ToLowerInvariant();
                if (baseName.StartsWith("infoproyect") && baseName.EndsWith(".txt"))
                {
                    return ParseInfoText(kv.Value);
                }
            }
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<string, string> ParseInfoText(byte[] bytes)
        {
            var info = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var text = Encoding.UTF8.GetString(bytes);
            foreach (var line in text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                var parts = line.Split(':', 2);
                if (parts.Length == 2)
                {
                    info[parts[0].Trim()] = parts[1].Trim();
                }
            }
            return info;
        }

        private static byte[]? FindLogoBytes(IDictionary<string, byte[]> archivos)
        {
            if (archivos != null)
            {
                foreach (var kv in archivos)
                {
                    var baseName = Path.GetFileName(kv.Key).ToLowerInvariant();
                    if (baseName.Contains("portada") && (baseName.EndsWith(".png") || baseName.EndsWith(".jpg") || baseName.EndsWith(".jpeg")))
                    {
                        return kv.Value;
                    }
                }
            }
            return null;
        }

        private static int ComputePercent(int done, int total)
        {
            if (total <= 0) return 100;
            return Math.Min(100, (int)Math.Round(done * 100.0 / total));
        }
    }
}
