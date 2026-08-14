using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Inforest.Application;
using Inforest.Infrastructure;
using Inforest.Desktop.POS;
using Inforest.Desktop.CajaRapida;
using Inforest.Desktop.Adicion;
using Inforest.Desktop.Caja;
using Inforest.Desktop.Motorizado;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop;

static class Program
{
    /// <summary>
    /// Punto de entrada de INFOREST Desktop .NET 8.
    /// Legacy: Sub Main() en modPuntoVenta.bas — inicialización de conexión y formulario principal.
    ///
    /// Bootstrap P3-03: formulario de acceso se resuelve desde DI
    /// para inicializar licencia, autenticación, sesión y auditoría.
    /// </summary>
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();

        using var host = CreateHostBuilder().Build();

        var appInstanceGuard = host.Services.GetRequiredService<IAppInstanceGuard>();
        if (!appInstanceGuard.TryAcquire("InfoRest.POS", out var instanceLease))
        {
            MessageBox.Show("Ya se está ejecutando el Aplicativo!", "Atención!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using (instanceLease)
        {
            // Formulario de acceso resuelto desde DI.
            var form = host.Services.GetRequiredService<Form1>();
            System.Windows.Forms.Application.Run(form);
        }
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

                // Formulario de acceso registrado en DI.
                services.AddTransient<Form1>();

                // Módulos principales — resolubles desde DI
                services.AddTransient<FrmPuntoVenta>();
                services.AddTransient<FrmCajaRapida>();
                services.AddTransient<FrmAdicion>();
                services.AddTransient<FrmCuadreCaja>();
                services.AddTransient<AsignacionMotorizadoForm>();
                services.AddTransient<ReasignacionMotorizadoForm>();
            });
    }
}