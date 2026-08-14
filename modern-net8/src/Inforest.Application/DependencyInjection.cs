using Microsoft.Extensions.DependencyInjection;
using Inforest.Application.Interfaces;
using Inforest.Application.Delivery;
using Inforest.Application.Despacho;
using Inforest.Application.Motorizado;
using Inforest.Application.CentralPedidos;
using Inforest.Application.Reportes;
using Inforest.Application.Configuracion;
using Inforest.Application.Maestros;
using Inforest.Application.Turno;
using Inforest.Application.Pedidos;
using Inforest.Application.Ventas;
using Inforest.Application.Caja;
using Inforest.Application.Seguridad;
using Inforest.Application.Impresion;
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

        // W1: Configuración — Handlers CQRS para TPARAMETRO/TCAJA (BR-SQL-003, BR-SQL-004)
        services.AddScoped<ObtenerConfiguracionSistemaHandler>();
        services.AddScoped<ObtenerConfiguracionCajaHandler>();
        services.AddScoped<ObtenerTodasCajasHandler>();
        services.AddScoped<ValidarInicioPosHandler>();
        // POS-FUNC-019: Seguridad / Contraseñas (BR-006)
        services.AddScoped<CambiarPasswordHandler>();

        // W2: Maestros — Handlers CQRS
        services.AddScoped<ObtenerGrupoProductoPorCodigoHandler>();
        services.AddScoped<ObtenerGrupoProductosActivosHandler>();
        services.AddScoped<CrearGrupoProductoHandler>();
        services.AddScoped<ActualizarGrupoProductoHandler>();
        services.AddScoped<ObtenerSubGrupoProductoPorCodigoHandler>();
        services.AddScoped<ObtenerSubGrupoProductosActivosHandler>();
        services.AddScoped<CrearSubGrupoProductoHandler>();
        services.AddScoped<ActualizarSubGrupoProductoHandler>();
        services.AddScoped<ObtenerClientePorCodigoHandler>();
        services.AddScoped<ObtenerClientesActivosHandler>();
        services.AddScoped<CrearClienteHandler>();
        services.AddScoped<ActualizarClienteHandler>();
        services.AddScoped<ObtenerMesaPorCodigoHandler>();
        services.AddScoped<ObtenerMesasActivosHandler>();
        services.AddScoped<CrearMesaHandler>();
        services.AddScoped<ActualizarMesaHandler>();
        services.AddScoped<ObtenerSalonPorCodigoHandler>();
        services.AddScoped<ObtenerSalonesActivosHandler>();
        services.AddScoped<CrearSalonHandler>();
        services.AddScoped<ActualizarSalonHandler>();
        services.AddScoped<ObtenerMozoPorCodigoHandler>();
        services.AddScoped<ObtenerMozosActivosHandler>();
        services.AddScoped<CrearMozoHandler>();
        services.AddScoped<ActualizarMozoHandler>();
        services.AddScoped<ObtenerProductoMaestroPorCodigoHandler>();
        services.AddScoped<ObtenerProductosMaestroActivosHandler>();
        services.AddScoped<CrearProductoMaestroHandler>();
        services.AddScoped<ActualizarProductoMaestroHandler>();

        // W3: Turno / Día Contable — Handlers CQRS (BR-004, BR-005)
        services.AddScoped<AbrirTurnoHandler>();
        services.AddScoped<CerrarTurnoHandler>();
        services.AddScoped<ObtenerTurnoActualHandler>();

        // W4: Pedidos — Handlers CQRS (BR-SQL-001, BR-SQL-002, BR-003)
        services.AddScoped<CreatePedidoHandler>();
        services.AddScoped<UpdatePedidoHandler>();
        services.AddScoped<ObtenerPedidoPorCodigoHandler>();
        services.AddScoped<ObtenerPedidosPorTurnoHandler>();
        services.AddScoped<ObtenerPedidosPorMesaHandler>();

        // W5: Venta / Documentos — Handlers CQRS (BR-002, BR-013)
        services.AddScoped<EmitirDocumentoHandler>();
        services.AddScoped<ObtenerDocumentoHandler>();
        services.AddScoped<AnularDocumentoHandler>();
        // POS-FUNC-006: Notas de Crédito (BR-NC-001..006)
        services.AddScoped<EmitirNotaCreditoHandler>();
        services.AddScoped<AnularNotaCreditoHandler>();
        services.AddScoped<ObtenerNotaCreditoHandler>();
        services.AddScoped<ObtenerNotasCreditoPorFechaHandler>();
        // POS-FUNC-007: Correlativos de comprobante (BR-POS-008)
        services.AddScoped<ObtenerCorrelativosPorCajaHandler>();
        services.AddScoped<ObtenerCorrelativoFacturaHandler>();
        services.AddScoped<ObtenerTodosCorrelativosHandler>();
        // POS-FUNC-016: Impresión pre-cuenta/comanda (BR-008)
        services.AddScoped<ImprimirPrecuentaHandler>();
        services.AddScoped<ObtenerImpresorasPorCajaHandler>();

        // W6: Caja / Pagos — Handlers CQRS (BR-007, BR-013)
        services.AddScoped<PagarDocumentoHandler>();
        services.AddScoped<ObtenerResumenTurnoHandler>();
        services.AddScoped<ObtenerMediosPagoHandler>();

        // P3-09: Delivery, Despacho, Motorizados y Centralización — Handlers
        services.AddScoped<CrearClienteDeliveryHandler>();
        services.AddScoped<ActualizarClienteDeliveryHandler>();
        services.AddScoped<ActualizarFotoClienteDeliveryHandler>();
        services.AddScoped<BuscarClienteDeliveryHandler>();
        services.AddScoped<ObtenerClientesFrecuentesHandler>();
        services.AddScoped<CrearPedidoDeliveryHandler>();
        services.AddScoped<ActualizarEstadoDeliveryHandler>();
        services.AddScoped<ObtenerDeliveryPendientesHandler>();
        services.AddScoped<ObtenerPedidosDespachadorHandler>();
        services.AddScoped<ExportarDespachadorHandler>();
        services.AddScoped<EmitirDespachoPedidoHandler>();
        services.AddScoped<ObtenerMotorizadosActivosHandler>();
        services.AddScoped<ObtenerPedidosReasignacionHandler>();
        services.AddScoped<ObtenerTarifaMotorizadoHandler>();
        services.AddScoped<AsignarMotorizadoHandler>();
        services.AddScoped<ReasignarMotorizadoHandler>();
        services.AddScoped<AsignarReasignacionMotorizadoHandler>();
        services.AddScoped<DesasignarReasignacionMotorizadoHandler>();
        services.AddScoped<RegistrarLlegadaSalidaHandler>();
        services.AddScoped<ActualizarTarifaMotorizadoHandler>();
        services.AddScoped<ObtenerPedidosCentralHandler>();
        services.AddScoped<ObtenerLocalesActivosHandler>();
        services.AddScoped<SincronizarPedidoCentralHandler>();
        services.AddScoped<ObtenerOrdenesExternasHandler>();
        services.AddScoped<EnviarDatosRappiHandler>();

        // P3-10: Reportes — Handlers (ADR-007)
        services.AddScoped<ObtenerReporteComandaHandler>();
        services.AddScoped<ObtenerReportePropinaHandler>();
        services.AddScoped<ObtenerReportePrincipalClienteHandler>();
        services.AddScoped<ObtenerReporteCtaCteIntegradoHandler>();
        services.AddScoped<ObtenerReportePaloteoComparativoHandler>();
        services.AddScoped<ObtenerReportePaloteoSubProdHandler>();
        services.AddScoped<ObtenerReportePaloteoVentaIntegradoHandler>();
        services.AddScoped<ObtenerReporteRankingIntegradoHandler>();
        services.AddScoped<ObtenerReporteVentaMensualIntegradoHandler>();
        services.AddScoped<ObtenerReporteAnaliticoMotorizadoIntegradoHandler>();
        services.AddScoped<ObtenerReporteTiempoKdsPedidoHandler>();
        services.AddScoped<ObtenerReporteTiempoKdsProductoHandler>();

        return services;
    }
}
