using Inforest.Application.Interfaces;
using Inforest.Application.Interfaces.Country;
using Inforest.Application.Interfaces.Hardware;
using Inforest.Application.Reportes;
using Inforest.Domain.Repositories;
using Inforest.Infrastructure.Country;
using Inforest.Infrastructure.Data;
using Inforest.Infrastructure.Delivery;
using Inforest.Infrastructure.Hardware;
using Inforest.Infrastructure.Kitchen;
using Inforest.Infrastructure.Motorizado;
using Inforest.Infrastructure.Reportes;
using Inforest.Infrastructure.Security;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace Inforest.Infrastructure;

/// <summary>
/// Configuración de DI para la capa Infrastructure.
/// </summary>
public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        // Configurar Serilog
        Log.Logger = new LoggerConfiguration()
            .ReadFrom.Configuration(configuration)
            .Enrich.FromLogContext()
            .WriteTo.Console()
            .WriteTo.File("logs/inforest-.log", rollingInterval: RollingInterval.Day)
            .CreateLogger();

        services.AddLogging(loggingBuilder =>
            loggingBuilder.AddSerilog(dispose: true));

        // P3-02: Contrato SQL y ejecución de datos Legacy
        services.AddSingleton<IDbConnectionFactory, SqlConnectionFactory>();
        services.AddSingleton<ISpExecutor, SpExecutor>();
        services.AddSingleton<IModernPasswordHashStore, ModernPasswordHashStore>();
        services.AddSingleton<ISessionService, SessionService>();
        services.AddScoped<IAuthService, AuthService>();
        services.AddSingleton<IRbacService, RbacService>();
        services.AddScoped<IAuditoriaService, AuditoriaService>();
        services.AddScoped<ILicenseService, LicenseService>();
        services.AddSingleton<IImpresoraService, NullImpresoraService>();
        services.AddScoped<IKitchenConfigurationProvider, KitchenConfigurationProvider>();
        services.AddScoped<IProduccionCocinaService, ProduccionCocinaService>();
        services.AddScoped<IKdsLegacyGateway, KdsLegacyGateway>();
        services.AddScoped<IComandaAreaDispatcher, ComandaAreaDispatcher>();
        services.AddScoped<IKdsDispatcher, KdsXmlDispatcher>();

        // P3-09: Delivery, Despacho, Motorizados y Centralización
        services.AddScoped<IClienteDeliveryRepository, ClienteDeliveryRepository>();
        services.AddScoped<IPedidoDeliveryRepository, PedidoDeliveryRepository>();
        services.AddScoped<IMotorizadoRepository, MotorizadoRepository>();
        services.AddScoped<ILocalRepository, LocalRepository>();
        services.AddScoped<ICentralPedidosRepository, CentralPedidosRepository>();
        services.AddScoped<IRappiOrderService, RappiOrderAdapter>();

        // Registrar repositorios aquí a medida que se implementen
        // services.AddScoped<IPedidoRepository, PedidoRepository>();
        // services.AddScoped<IProductoRepository, ProductoRepository>();

        // P3-10: Reportes — IReporteRepository + ReporteRepository (ADR-007)
        services.AddScoped<IReporteRepository, ReporteRepository>();

        // P3-11: Periféricos especializados e integraciones por país
        services.AddHardwareServices(configuration);
        services.AddCountryServices(configuration);

        return services;
    }

    private static void AddHardwareServices(this IServiceCollection services, IConfiguration configuration)
    {
        // PinPad — BR-PERIPH-001, ADR-009
        services.Configure<PinPadOptions>(configuration.GetSection(PinPadOptions.SectionName));
        var pinPadHabilitado = configuration.GetValue<bool>($"{PinPadOptions.SectionName}:Habilitado");
        if (pinPadHabilitado)
            services.AddSingleton<IPinPadService, PinPadService>();
        else
            services.AddSingleton<IPinPadService, NullPinPadService>();

        // CashDro — BR-PERIPH-002, ADR-011
        services.Configure<CashDroOptions>(configuration.GetSection(CashDroOptions.SectionName));
        var cashDroHabilitado = configuration.GetValue<bool>($"{CashDroOptions.SectionName}:Habilitado");
        if (cashDroHabilitado)
            services.AddSingleton<ICashDroService, CashDroService>();
        else
            services.AddSingleton<ICashDroService, NullCashDroService>();

        // BlueVision — BR-PERIPH-003, ADR-012
        services.Configure<BlueVisionOptions>(configuration.GetSection(BlueVisionOptions.SectionName));
        var blueVisionHabilitado = configuration.GetValue<bool>($"{BlueVisionOptions.SectionName}:Habilitado");
        if (blueVisionHabilitado)
        {
            services.AddHttpClient<IBlueVisionService, BlueVisionHttpClient>(client =>
            {
                var url = configuration.GetValue<string>($"{BlueVisionOptions.SectionName}:Url") ?? "http://localhost/";
                client.BaseAddress = new Uri(url);
            });
        }
        else
        {
            services.AddSingleton<IBlueVisionService, NullBlueVisionService>();
        }

        // Biometría SecuGen — BR-PERIPH-004, ADR-009 (GAP: OCX 32-bit sin SDK .NET)
        services.Configure<BiometriaOptions>(configuration.GetSection(BiometriaOptions.SectionName));
        services.AddSingleton<IBiometriaService, NullBiometriaService>();

        // Impresora Fiscal Epson — BR-PERIPH-005, ADR-009 (GAP: OCX 32-bit sin SDK .NET)
        services.Configure<ImpresoraFiscalOptions>(configuration.GetSection(ImpresoraFiscalOptions.SectionName));
        services.AddSingleton<IImpresoraFiscalService, NullImpresoraFiscalService>();
    }

    private static void AddCountryServices(this IServiceCollection services, IConfiguration configuration)
    {
        // FE Options — BR-PERIPH-006, ADR-010
        services.Configure<PeruFEOptions>(configuration.GetSection(PeruFEOptions.SectionName));
        services.Configure<ArgentinaFEOptions>(configuration.GetSection(ArgentinaFEOptions.SectionName));

        // Gateways de FE por país (todos registrados para resolución por fábrica)
        services.AddSingleton<IFacturacionElectronicaGateway, NullFEGateway>();
        var feHabilitado = configuration.GetValue<bool>("FE:Habilitado");
        if (feHabilitado)
        {
            services.AddSingleton<IFacturacionElectronicaGateway, PeruFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, ArgentinaFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, ChileFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, EcuadorFEGateway>();
            services.AddSingleton<IFacturacionElectronicaGateway, BoliviaFEGateway>();
        }
        services.AddSingleton<FacturacionElectronicaFactory>();

        // Políticas de país — BR-PERIPH-007, ADR-008
        services.AddSingleton<IPaisPolicy, PeruPaisPolicy>();
        services.AddSingleton<IPaisPolicy, ArgentinaPaisPolicy>();
        services.AddSingleton<IPaisPolicy, ChilePaisPolicy>();
        services.AddSingleton<IPaisPolicy, EcuadorPaisPolicy>();
        services.AddSingleton<IPaisPolicy, BoliviaPaisPolicy>();
        services.AddSingleton<IPaisPolicy, EspanaPaisPolicy>();
        services.AddSingleton<PaisPolicyFactory>();
    }
}
