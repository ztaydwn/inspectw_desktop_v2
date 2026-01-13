using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using InspectW.Domain;
using InspectW.Infrastructure;
using InspectW.Reporting;
using Microsoft.Win32;

namespace InspectW.App.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly IZipLoader _zipLoader;
        private readonly IFolderLoader _folderLoader;
        private readonly IGroupingService _groupingService;
        private readonly IControlDocumentLoader _controlLoader;
        private readonly IXlsxReportService _xlsxReportService;

        private IDictionary<string, byte[]> _lastArchivos = new Dictionary<string, byte[]>();

        [ObservableProperty] private string? _zipPath;
        [ObservableProperty] private string? _folderPath;
        [ObservableProperty] private string? _historicoPath;
        [ObservableProperty] private string _status = "Listo";
        [ObservableProperty] private bool _isBusy;

        public ObservableCollection<Grupo> Grupos { get; } = new();

        public IAsyncRelayCommand LoadCommand { get; }
        public IAsyncRelayCommand ExportXlsxCommand { get; }
        public ICommand BrowseZipCommand { get; }
        public ICommand BrowseFolderCommand { get; }
        public ICommand BrowseHistoricoCommand { get; }

        public MainViewModel(
            IZipLoader zipLoader,
            IFolderLoader folderLoader,
            IGroupingService groupingService,
            IControlDocumentLoader controlLoader,
            IXlsxReportService xlsxReportService)
        {
            _zipLoader = zipLoader;
            _folderLoader = folderLoader;
            _groupingService = groupingService;
            _controlLoader = controlLoader;
            _xlsxReportService = xlsxReportService;

            LoadCommand = new AsyncRelayCommand(LoadAsync, CanRun);
            ExportXlsxCommand = new AsyncRelayCommand(ExportAsync, () => Grupos.Count > 0 && !IsBusy);
            BrowseZipCommand = new RelayCommand(BrowseZip);
            BrowseFolderCommand = new RelayCommand(BrowseFolder);
            BrowseHistoricoCommand = new RelayCommand(BrowseHistorico);
        }

        private bool CanRun() => !IsBusy;

        private void BrowseZip()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "ZIP|*.zip",
                CheckFileExists = true
            };
            if (dlg.ShowDialog() == true)
            {
                ZipPath = dlg.FileName;
                FolderPath = null;
            }
        }

        private void BrowseFolder()
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FolderPath = dlg.SelectedPath;
                ZipPath = null;
            }
        }

        private void BrowseHistorico()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "CSV|*.csv",
                CheckFileExists = true
            };
            if (dlg.ShowDialog() == true)
            {
                HistoricoPath = dlg.FileName;
            }
        }

        private async Task LoadAsync()
        {
            if (IsBusy) return;
            IsBusy = true;
            Status = "Cargando...";
            Grupos.Clear();
            _lastArchivos = new Dictionary<string, byte[]>();

            try
            {
                IDictionary<string, byte[]> archivos;
                if (!string.IsNullOrWhiteSpace(ZipPath))
                {
                    archivos = await _zipLoader.LoadAsync(ZipPath!);
                }
                else if (!string.IsNullOrWhiteSpace(FolderPath))
                {
                    archivos = await _folderLoader.LoadAsync(FolderPath!);
                }
                else
                {
                    Status = "Selecciona ZIP o carpeta.";
                    return;
                }

                var historicoPath = HistoricoPath;
                if (string.IsNullOrWhiteSpace(historicoPath))
                {
                    var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    var candidate = Path.Combine(baseDir, "datos", "historico.csv");
                    if (File.Exists(candidate))
                    {
                        historicoPath = candidate;
                    }
                }
                if (!string.IsNullOrWhiteSpace(historicoPath) && File.Exists(historicoPath))
                {
                    archivos["historico.csv"] = await File.ReadAllBytesAsync(historicoPath);
                }

                _lastArchivos = archivos;

                var (grupos, errores) = await _groupingService.AgruparAsync(archivos);
                foreach (var g in grupos)
                {
                    Grupos.Add(g);
                }

                Status = errores.Count > 0 ? string.Join(" | ", errores) : $"Cargado {Grupos.Count} grupos.";
            }
            catch (Exception ex)
            {
                Status = $"Error: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                LoadCommand.NotifyCanExecuteChanged();
                ExportXlsxCommand.NotifyCanExecuteChanged();
            }
        }

        private async Task ExportAsync()
        {
            if (IsBusy || Grupos.Count == 0) return;

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = "reporte.xlsx"
            };
            if (dlg.ShowDialog() != true)
                return;

            IsBusy = true;
            Status = "Generando XLSX...";

            try
            {
                var controlDocs = _controlLoader.Load(_lastArchivos);

                await _xlsxReportService.GenerateAsync(
                    Grupos.ToList(),
                    _lastArchivos,
                    dlg.FileName,
                    null,
                    controlDocs,
                    progreso: null);

                Status = $"Exportado: {dlg.FileName}";
            }
            catch (Exception ex)
            {
                Status = $"Error exportando: {ex.Message}";
            }
            finally
            {
                IsBusy = false;
                ExportXlsxCommand.NotifyCanExecuteChanged();
                LoadCommand.NotifyCanExecuteChanged();
            }
        }
    }
}
