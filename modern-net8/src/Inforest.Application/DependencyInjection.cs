using Microsoft.Extensions.DependencyInjection;
using Inforest.Application.Interfaces;
using Inforest.Domain.Services;

namespace Inforest.Application;

/// <summary>
/// Configuración de DI para la capa Application.
/// </summary>
public static class DependencyInjection
{
    public static IServiceCollection AddApplication(this IServiceCollection services)
    {
        services.AddSingleton<AreaProduccionRouter>();
        return services;
    }
}
