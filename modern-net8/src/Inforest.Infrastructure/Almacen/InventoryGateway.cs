using Dapper;
using Inforest.Application.Interfaces;
using Microsoft.Extensions.Logging;

namespace Inforest.Infrastructure.Almacen;

/// <summary>
/// Pasarela de almacén que ejecuta descargos sobre la base ALMACEN.
/// Legacy: clsAlmacen.cls — BR-008.
/// SPs: usp_Inforest_DescargoVenta, usp_Inforest_DescargoVentaPorInsumo, USP_MODIFICARINSUMOS, usp_Inforest_InsertaLogDescargo.
/// </summary>
internal sealed class InventoryGateway : IInventoryGateway
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;
    private readonly ILogger<InventoryGateway> _logger;

    public InventoryGateway(
        IDbConnectionFactory connectionFactory,
        ISpExecutor spExecutor,
        ILogger<InventoryGateway> logger)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
        _logger = logger;
    }

    /// <inheritdoc/>
    public async Task<bool> EjecutarDescargoVentaAsync(
        string codigoAlmacen,
        DateTime fechaInicio,
        DateTime fechaFin,
        string tablaTemporal,
        string codigoLocal,
        string codigoPedido,
        int tipoOperacion,
        CancellationToken ct = default)
    {
        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
            await _spExecutor.ExecuteAsync(
                connection,
                "usp_Inforest_DescargoVenta",
                new
                {
                    Almacen   = codigoAlmacen,
                    fechaIni  = fechaInicio.ToString("yyyy-MM-dd"),
                    fechaFin  = fechaFin.ToString("yyyy-MM-dd"),
                    sTemporal = tablaTemporal,
                    Local     = codigoLocal,
                    Pedido    = codigoPedido,
                    tipooper  = tipoOperacion
                },
                cancellationToken: ct);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error ejecutando descargo venta para pedido {Pedido}", codigoPedido);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task<bool> EjecutarDescargoVentaPorInsumoAsync(
        string codigoPedido,
        string codigoAlmacen,
        DateTime fechaDiaContable,
        string usuario,
        CancellationToken ct = default)
    {
        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
            await _spExecutor.ExecuteAsync(
                connection,
                "usp_Inforest_DescargoVentaPorInsumo",
                new
                {
                    tPedido      = codigoPedido,
                    tAlmacen     = codigoAlmacen,
                    fDiaContable = fechaDiaContable.Date,
                    tUsuario     = usuario
                },
                cancellationToken: ct);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error ejecutando descargo por insumo para pedido {Pedido}", codigoPedido);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task<bool> ModificarInsumosAsync(
        string codigoPedido,
        string codigoAlmacen,
        string usuario,
        CancellationToken ct = default)
    {
        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
            await _spExecutor.ExecuteAsync(
                connection,
                "USP_MODIFICARINSUMOS",
                new
                {
                    tPedido  = codigoPedido,
                    tAlmacen = codigoAlmacen,
                    tUsuario = usuario
                },
                cancellationToken: ct);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error modificando insumos para pedido {Pedido}", codigoPedido);
            return false;
        }
    }

    /// <inheritdoc/>
    public async Task InsertarLogDescargoAsync(
        int correlativo,
        DateTime fechaRegistro,
        string usuario,
        DateTime fechaInicio,
        DateTime fechaFin,
        int cantidadPedidos,
        int cantidadPedidosOk,
        int cantidadPedidosError,
        string equipoIp,
        string equipoNombre,
        string equipoUsuario,
        CancellationToken ct = default)
    {
        try
        {
            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
            await _spExecutor.ExecuteAsync(
                connection,
                "usp_Inforest_InsertaLogDescargo",
                new
                {
                    num_nCorrelativo  = correlativo,
                    dtm_fRegInicio    = fechaRegistro,
                    vch_tUsuario      = usuario,
                    dtm_fFecInicio    = fechaInicio,
                    dtm_fFecFin       = fechaFin,
                    num_nCantPedidos  = cantidadPedidos,
                    num_nCantPedOk    = cantidadPedidosOk,
                    num_nCantPedError = cantidadPedidosError,
                    vch_tEquIp        = equipoIp,
                    vch_tEquName      = equipoNombre,
                    vch_tEquUser      = equipoUsuario
                },
                cancellationToken: ct);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error insertando log de descargo correlativo {Correlativo}", correlativo);
        }
    }
}
