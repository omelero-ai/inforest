using Microsoft.Extensions.DependencyInjection;
using Inforest.Application.Interfaces;
using Inforest.Application.Delivery;
using Inforest.Application.Despacho;
using Inforest.Application.Motorizado;
using Inforest.Application.CentralPedidos;
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

        // P3-09: Delivery, Despacho, Motorizados y Centralización — Handlers
        services.AddScoped<CrearClienteDeliveryHandler>();
        services.AddScoped<ActualizarClienteDeliveryHandler>();
        services.AddScoped<BuscarClienteDeliveryHandler>();
        services.AddScoped<ObtenerClientesFrecuentesHandler>();
        services.AddScoped<CrearPedidoDeliveryHandler>();
        services.AddScoped<ActualizarEstadoDeliveryHandler>();
        services.AddScoped<ObtenerDeliveryPendientesHandler>();
        services.AddScoped<ObtenerPedidosDespachadorHandler>();
        services.AddScoped<ExportarDespachadorHandler>();
        services.AddScoped<EmitirDespachoPedidoHandler>();
        services.AddScoped<ObtenerMotorizadosActivosHandler>();
        services.AddScoped<ObtenerTarifaMotorizadoHandler>();
        services.AddScoped<AsignarMotorizadoHandler>();
        services.AddScoped<ReasignarMotorizadoHandler>();
        services.AddScoped<RegistrarLlegadaSalidaHandler>();
        services.AddScoped<ActualizarTarifaMotorizadoHandler>();
        services.AddScoped<ObtenerPedidosCentralHandler>();
        services.AddScoped<ObtenerLocalesActivosHandler>();
        services.AddScoped<SincronizarPedidoCentralHandler>();
        services.AddScoped<ObtenerOrdenesExternasHandler>();
        services.AddScoped<EnviarDatosRappiHandler>();

        return services;
    }
}
