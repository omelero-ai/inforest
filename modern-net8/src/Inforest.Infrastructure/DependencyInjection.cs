using Inforest.Application.Interfaces;
using Inforest.Infrastructure.Data;
using Inforest.Infrastructure.Kitchen;
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

        // Registrar repositorios aquí a medida que se implementen
        // services.AddScoped<IPedidoRepository, PedidoRepository>();
        // services.AddScoped<IProductoRepository, ProductoRepository>();

        return services;
    }
}
