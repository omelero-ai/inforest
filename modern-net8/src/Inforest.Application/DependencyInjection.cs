using Microsoft.Extensions.DependencyInjection;
using Inforest.Application.Interfaces;

namespace Inforest.Application;

/// <summary>
/// Configuración de DI para la capa Application.
/// </summary>
public static class DependencyInjection
{
    public static IServiceCollection AddApplication(this IServiceCollection services)
    {
        // Registrar servicios de aplicación aquí a medida que se migren
        return services;
    }
}
