using Inforest.Application.Interfaces;
using Inforest.Application.Interfaces.Country;
using Inforest.Application.Interfaces.Hardware;
using Inforest.Application.Reportes;
using Inforest.Infrastructure;
using Inforest.Infrastructure.Country;
using Inforest.Infrastructure.Hardware;
using Inforest.Infrastructure.Kitchen;
using Inforest.Infrastructure.Reportes;
using Inforest.Infrastructure.Security;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace Inforest.Infrastructure.Tests.DependencyInjection;

/// <summary>
/// Validación integral de wiring transversal de Fase 3.
/// Legacy: modPuntoVenta.bas, modKDS.bas, DLL3500.bas, modBlueVision.bas, modProcedimientoNuevo.bas.
/// </summary>
public sealed class InfrastructureDependencyInjectionTests
{
    [Fact]
    public void AddInfrastructure_ConPerifericosDeshabilitados_ResuelveServiciosNullYCore()
    {
        var services = new ServiceCollection();
        var configuration = BuildConfiguration();

        services.AddSingleton<IConfiguration>(configuration);
        services.AddInfrastructure(configuration);

        using var provider = services.BuildServiceProvider();

        Assert.IsType<NullPinPadService>(provider.GetRequiredService<IPinPadService>());
        Assert.IsType<NullCashDroService>(provider.GetRequiredService<ICashDroService>());
        Assert.IsType<NullBlueVisionService>(provider.GetRequiredService<IBlueVisionService>());
        Assert.IsType<NullBiometriaService>(provider.GetRequiredService<IBiometriaService>());
        Assert.IsType<NullImpresoraFiscalService>(provider.GetRequiredService<IImpresoraFiscalService>());
        Assert.IsType<SessionService>(provider.GetRequiredService<ISessionService>());
        Assert.IsType<ReporteRepository>(provider.GetRequiredService<IReporteRepository>());
        Assert.IsType<KdsXmlDispatcher>(provider.GetRequiredService<IKdsDispatcher>());
    }

    [Fact]
    public void AddInfrastructure_ConPerifericosYFeHabilitados_ResuelveImplementacionesConcretas()
    {
        var services = new ServiceCollection();
        var configuration = BuildConfiguration(new Dictionary<string, string?>
        {
            ["Hardware:PinPad:Habilitado"] = "true",
            ["Hardware:PinPad:DllPath"] = "/tmp/caja_pinpad.dll",
            ["Hardware:CashDro:Habilitado"] = "true",
            ["Hardware:CashDro:RutaMotor"] = "/tmp/MotorCashDrow.exe",
            ["Hardware:BlueVision:Habilitado"] = "true",
            ["Hardware:BlueVision:Url"] = "http://localhost:5010/",
            ["Hardware:BlueVision:Login"] = "demo",
            ["Hardware:BlueVision:ClearPassword"] = "demo",
            ["FE:Habilitado"] = "true"
        });

        services.AddSingleton<IConfiguration>(configuration);
        services.AddInfrastructure(configuration);

        using var provider = services.BuildServiceProvider();

        Assert.IsType<PinPadService>(provider.GetRequiredService<IPinPadService>());
        Assert.IsType<CashDroService>(provider.GetRequiredService<ICashDroService>());
        Assert.IsType<BlueVisionHttpClient>(provider.GetRequiredService<IBlueVisionService>());

        var gateways = provider.GetServices<IFacturacionElectronicaGateway>().Select(g => g.CodigoPais).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var codigoPais in new[] { "NULL", "PE", "AR", "CL", "EC", "BO" })
            Assert.Contains(codigoPais, gateways);

        var factory = provider.GetRequiredService<FacturacionElectronicaFactory>();
        Assert.Equal("PE", factory.ObtenerGateway("PE").CodigoPais);

        var paisFactory = provider.GetRequiredService<PaisPolicyFactory>();
        Assert.Equal("AR", paisFactory.ObtenerPolitica("AR").CodigoPais);
    }

    private static IConfiguration BuildConfiguration(Dictionary<string, string?>? overrides = null)
    {
        var values = new Dictionary<string, string?>
        {
            ["ConnectionStrings:Inforest"] = "Server=(local);Database=INFOREST;Trusted_Connection=true;",
            ["Hardware:PinPad:Habilitado"] = "false",
            ["Hardware:CashDro:Habilitado"] = "false",
            ["Hardware:BlueVision:Habilitado"] = "false",
            ["FE:Habilitado"] = "false"
        };

        if (overrides is not null)
        {
            foreach (var pair in overrides)
                values[pair.Key] = pair.Value;
        }

        return new ConfigurationBuilder()
            .AddInMemoryCollection(values)
            .Build();
    }
}
