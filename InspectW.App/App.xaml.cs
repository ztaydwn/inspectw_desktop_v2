using System;
using System.Windows;
using InspectW.Infrastructure;
using InspectW.Reporting;
using Microsoft.Extensions.DependencyInjection;
using InspectW.App.ViewModels;
using Application = System.Windows.Application;

namespace InspectW.App;

public partial class App : Application
{
    public IServiceProvider Services { get; private set; } = null!;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        Services = ConfigureServices();
        var mainWindow = new MainWindow();
        mainWindow.Show();
    }

    private static IServiceProvider ConfigureServices()
    {
        var services = new ServiceCollection();

        services.AddSingleton<IZipLoader, ZipLoader>();
        services.AddSingleton<IFolderLoader, FolderLoader>();
        services.AddSingleton<IDescriptionParser, DescriptionParser>();
        services.AddSingleton<IControlDocumentLoader, ControlDocumentLoader>();
        services.AddSingleton<IRecommender>(sp =>
        {
            // se inicializa con CSV por archivo, el motor se construye por cada Apply usando historico*.csv de archivos
            return RecommendationEngine.FromCsv(Array.Empty<byte>());
        });
        services.AddSingleton<IGroupingService, GroupingService>();
        services.AddSingleton<IXlsxReportService, XlsxReportService>();
        services.AddSingleton<ViewModels.MainViewModel>();

        return services.BuildServiceProvider();
    }
}
