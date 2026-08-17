using Microsoft.Extensions.DependencyInjection;
using Inforest.Application.Interfaces;
using Inforest.Application.Almacen;
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
using Inforest.Application.Reservas;
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
        services.AddScoped<ObtenerPedidosSinMesaHandler>();

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
        // POS-FUNC-016: Impresión pre-cuenta/comanda (BR-PRECUENTA-001..003)
        services.AddScoped<ImprimirPrecuentaHandler>();
        services.AddScoped<ObtenerImpresorasPorCajaHandler>();

        // POS-FUNC-011: Reservas (BR-RESERVA-001..004)
        services.AddScoped<CrearReservaHandler>();
        services.AddScoped<ModificarReservaHandler>();
        services.AddScoped<AnularReservaHandler>();
        services.AddScoped<ObtenerReservaHandler>();
        services.AddScoped<ObtenerReservasPorFechaHandler>();
        services.AddScoped<ConvertirReservaAPedidoHandler>();

        // W6: Caja / Pagos — Handlers CQRS (BR-007, BR-013)
        services.AddScoped<PagarDocumentoHandler>();
        services.AddScoped<RegistrarPagosMultiplesHandler>();
        services.AddScoped<ObtenerResumenTurnoHandler>();
        services.AddScoped<ObtenerMediosPagoHandler>();
        services.AddScoped<ObtenerTerminalesPinPadHandler>();
        services.AddScoped<ProcesarPagoPinPadHandler>();

        // P3-09: Delivery, Despacho, Motorizados y Centralización — Handlers
        services.AddScoped<CrearClienteDeliveryHandler>();
        services.AddScoped<ActualizarClienteDeliveryHandler>();
        services.AddScoped<ActualizarFotoClienteDeliveryHandler>();
        services.AddScoped<BuscarClienteDeliveryHandler>();
        services.AddScoped<ObtenerClientesFrecuentesHandler>();
        services.AddScoped<ObtenerTarjetasProximidadHandler>();
        services.AddScoped<ObtenerMovimientosTarjetaProximidadHandler>();
        services.AddScoped<CrearTarjetaProximidadHandler>();
        services.AddScoped<ActualizarTarjetaProximidadHandler>();
        services.AddScoped<ObtenerRecargasTarjetaHandler>();
        services.AddScoped<RegistrarRecargaTarjetaHandler>();
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
        services.AddScoped<ConfirmarEntregaCentralHandler>();
        services.AddScoped<RevertirEntregaCentralHandler>();
        services.AddScoped<ModificarFechaProgramadaDeliveryHandler>();
        services.AddScoped<ObtenerPedidosSeguimientoDeliveryHandler>();

        // POS-FUNC-014: Importación de pedidos externos desde ALMACEN
        services.AddScoped<ObtenerRequerimientosPendientesHandler>();
        services.AddScoped<ObtenerDetalleRequerimientoHandler>();
        services.AddScoped<ImportarRequerimientoHandler>();

        // P3-10: Reportes — Handlers (ADR-007)
        services.AddScoped<ObtenerReporteComandaHandler>();
        services.AddScoped<ObtenerReportePropinaHandler>();
        services.AddScoped<ObtenerReportePrincipalClienteHandler>();
        services.AddScoped<ObtenerReporteCtaCteIntegradoHandler>();
        services.AddScoped<ObtenerReporteCtaCteOperativaHandler>();
        services.AddScoped<ObtenerTiposCtaCteReporteHandler>();
        services.AddScoped<ObtenerSubTiposCtaCteReporteHandler>();
        services.AddScoped<ObtenerReportePaloteoComparativoHandler>();
        services.AddScoped<ObtenerReportePaloteoSubProdHandler>();
        services.AddScoped<ObtenerReportePaloteoVentaIntegradoHandler>();
        services.AddScoped<ObtenerReporteRankingIntegradoHandler>();
        services.AddScoped<ObtenerReporteVentaMensualIntegradoHandler>();
        services.AddScoped<ObtenerReporteAnaliticoMotorizadoIntegradoHandler>();
        services.AddScoped<ObtenerReporteTiempoKdsPedidoHandler>();
        services.AddScoped<ObtenerReporteTiempoKdsProductoHandler>();
        services.AddScoped<ObtenerReporteAnulacionHandler>();
        services.AddScoped<ObtenerReporteLiquidacionTicketHandler>();
        services.AddScoped<ObtenerReportePaloteoTicketHandler>();
        services.AddScoped<ObtenerReporteEntregaHandler>();
        services.AddScoped<ObtenerReporteRegistroVentaHandler>(); // BR-REP-022

        // POS-FUNC-018: Recibo Ingresos/Egresos — MINGRESO / MEGRESO
        services.AddScoped<ObtenerIngresosHandler>();
        services.AddScoped<RegistrarIngresoHandler>();
        services.AddScoped<AnularIngresoHandler>();
        services.AddScoped<ObtenerEgresosHandler>();
        services.AddScoped<RegistrarEgresoHandler>();
        services.AddScoped<AnularEgresoHandler>();

        // POS-FUNC-029: Lista Motivos Anulación — vMotivoAnulacion (BR-MOTIVO-001..003)
        services.AddScoped<ObtenerMotivosAnulacionHandler>();
        // POS-FUNC-030: Solicitud de boleta con/sin datos (frmSolicitudBoleta.frm)
        services.AddScoped<ObtenerClienteGeneralBoletaHandler>();

        return services;
    }
}
