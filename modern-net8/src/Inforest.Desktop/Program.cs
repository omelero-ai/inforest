using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Inforest.Application;
using Inforest.Infrastructure;

namespace Inforest.Desktop;

static class Program
{
    /// <summary>
    /// Punto de entrada de INFOREST Desktop .NET 8.
    /// Legacy: Sub Main() en modPuntoVenta.bas — inicialización de conexión y formulario principal.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        var host = CreateHostBuilder().Build();

        var services = host.Services;

        // TODO: Reemplazar Form1 con formulario principal real (equivalente a frmPrincipal.frm)
        Application.Run(new Form1());
    }

    static IHostBuilder CreateHostBuilder()
    {
        return Host.CreateDefaultBuilder()
            .ConfigureAppConfiguration((context, config) =>
            {
                config.SetBasePath(AppContext.BaseDirectory);
                config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                config.AddJsonFile($"appsettings.{context.HostingEnvironment.EnvironmentName}.json",
                    optional: true, reloadOnChange: true);
                config.AddEnvironmentVariables("INFOREST_");
            })
            .ConfigureServices((context, services) =>
            {
                services.AddApplication();
                services.AddInfrastructure(context.Configuration);
            });
    }
}