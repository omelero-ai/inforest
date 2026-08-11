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
    ///
    /// Bootstrap corregido P3-01: formulario principal se resuelve desde el contenedor DI
    /// para que las capas Application e Infrastructure sean inyectables en la UI.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        var host = CreateHostBuilder().Build();

        // Formulario principal resuelto desde DI — permite inyección de servicios de Application.
        // TODO: Reemplazar Form1 con formulario principal real (equivalente a frmPrincipal.frm).
        var form = host.Services.GetRequiredService<Form1>();
        System.Windows.Forms.Application.Run(form);
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

                // Formularios registrados en DI para recibir servicios inyectados.
                // TODO: Reemplazar Form1 con formulario principal real.
                services.AddTransient<Form1>();
            });
    }
}