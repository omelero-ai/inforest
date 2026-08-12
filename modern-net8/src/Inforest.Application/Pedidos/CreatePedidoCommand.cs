using System.Data;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Cocina;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Pedidos;

/// <summary>
/// Legacy: spIns_MPEDIDO. BR-SQL-001.
/// </summary>
public record CreatePedidoCommand(
    string CodigoCaja,
    string CodigoTurno,
    string CodigoMesa,
    string CodigoMozo,
    string CodigoUsuario,
    CanalVenta Canal,
    int NumeroAdultos,
    int NumeroNinos,
    string? Observacion,
    IReadOnlyList<CreateDetalleItem> Items
);

public record CreateDetalleItem(
    string CodigoProducto,
    string CodigoGrupo,
    string CodigoSubGrupo,
    decimal Cantidad,
    decimal PrecioUnitario,
    string? Area,
    string? Observacion
);

public class CreatePedidoHandler
{
    private readonly IPedidoRepository _pedidoRepository;
    private readonly ITurnoRepository _turnoRepository;
    private readonly IProduccionCocinaService? _produccionCocinaService;
    private readonly IDbConnectionFactory? _connectionFactory;
    private readonly ISpExecutor? _spExecutor;

    public CreatePedidoHandler(
        IPedidoRepository pedidoRepository,
        ITurnoRepository turnoRepository,
        IProduccionCocinaService? produccionCocinaService = null,
        IDbConnectionFactory? connectionFactory = null,
        ISpExecutor? spExecutor = null)
    {
        _pedidoRepository = pedidoRepository;
        _turnoRepository = turnoRepository;
        _produccionCocinaService = produccionCocinaService;
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<Result<Pedido>> HandleAsync(CreatePedidoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoTurno))
            return Result.Fail<Pedido>("El turno es obligatorio.", "PEDIDO_TURNO_REQUERIDO");

        var turnoActual = await _turnoRepository.ObtenerTurnoActualAsync(command.CodigoCaja, ct);
        if (turnoActual is null || !string.Equals(turnoActual.CodigoTurno, command.CodigoTurno, StringComparison.OrdinalIgnoreCase))
            return Result.Fail<Pedido>("No existe un turno activo válido para la caja indicada.", "PEDIDO_TURNO_INVALIDO");

        var pedido = Pedido.Crear(
            GenerarCodigoTemporal(),
            command.CodigoMesa,
            command.CodigoTurno,
            command.CodigoUsuario,
            command.Canal,
            Math.Max(0, command.NumeroAdultos + command.NumeroNinos));

        PedidoPrivateAccessor.SetObservacion(pedido, command.Observacion);
        PedidoPrivateAccessor.SetFechaDiaContable(pedido, turnoActual.FechaDiaContable);

        foreach (var item in command.Items)
        {
            var detalle = DetallePedido.Crear(
                pedido.CodigoPedido,
                SiguienteItem(pedido.Detalles.Count),
                item.CodigoProducto,
                item.Cantidad,
                item.PrecioUnitario,
                item.PrecioUnitario);

            PedidoPrivateAccessor.ConfigurarDetalle(detalle, item.CodigoGrupo, item.CodigoSubGrupo, item.Area, item.Observacion);
            pedido.AgregarDetalle(detalle);
        }

        PedidoRepositoryContext.Set(new PedidoRepositoryMetadata(
            command.CodigoCaja,
            command.CodigoTurno,
            command.CodigoMesa,
            command.CodigoMozo,
            command.CodigoUsuario,
            command.Canal,
            command.NumeroAdultos,
            command.NumeroNinos,
            command.Observacion,
            turnoActual.FechaDiaContable));

        try
        {
            await _pedidoRepository.InsertarAsync(pedido, ct);
        }
        finally
        {
            PedidoRepositoryContext.Clear();
        }

        if (command.Items.Count > 0)
        {
            if (_connectionFactory is null || _spExecutor is null)
                return Result.Fail<Pedido>("No se configuró el acceso a BD para registrar el detalle del pedido.", "PEDIDO_DETALLE_DB_NO_CONFIGURADO");

            using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
            for (var index = 0; index < command.Items.Count; index++)
            {
                var item = command.Items[index];
                await _spExecutor.ExecuteAsync(
                    connection,
                    "spIns_DPEDIDO",
                    new
                    {
                        tPedido = pedido.CodigoPedido,
                        tTipoPedido = LegacyMappings.MapCanal(command.Canal),
                        tItem = SiguienteItem(index),
                        tProducto = item.CodigoProducto,
                        tGrupo = item.CodigoGrupo,
                        tSubGrupo = item.CodigoSubGrupo,
                        nPBase = item.PrecioUnitario,
                        nRecargo = 0m,
                        nDescuento = 0m,
                        nOficial = item.PrecioUnitario,
                        nImpuesto1 = 0m,
                        nImpuesto2 = 0m,
                        nImpuesto3 = 0m,
                        nPVenta = item.PrecioUnitario,
                        nCantidad = item.Cantidad,
                        tArea = item.Area ?? string.Empty,
                        lImprimeArea = !string.IsNullOrWhiteSpace(item.Area),
                        lCombinacion = false,
                        nCombinacion = 0m,
                        tComanda = string.Empty,
                        tMozo = command.CodigoMozo,
                        tUsuario = command.CodigoUsuario,
                        nInsumo = 0m,
                        nGasto = 0m,
                        nMObra = 0m,
                        nOrden = index + 1,
                        tUnidadNegocio = string.Empty,
                        tOferta = string.Empty,
                        tsubalmacen = string.Empty,
                        tFechaContable = turnoActual.FechaDiaContable,
                        sCaja = command.CodigoCaja
                    },
                    cancellationToken: ct);
            }

            if (_produccionCocinaService is not null)
            {
                var orden = OrdenProduccion.Crear(
                    pedido.CodigoPedido,
                    command.CodigoCaja,
                    command.CodigoMesa,
                    command.CodigoMesa,
                    null,
                    Environment.MachineName,
                    command.CodigoUsuario,
                    command.Items.Select((item, index) =>
                        ItemProduccion.Crear(
                            pedido.CodigoPedido,
                            SiguienteItem(index),
                            item.CodigoProducto,
                            item.CodigoProducto,
                            item.Cantidad,
                            item.Area,
                            !string.IsNullOrWhiteSpace(item.Area))));

                var configuracion = ConfiguracionProduccionCocina.Crear(false, false, false, false, null, null, null, null);
                var result = await _produccionCocinaService.EnviarNuevaOrdenAsync(orden, configuracion, ct);
                if (!result.EsExitoso)
                    return Result.Fail<Pedido>(result.MensajeError ?? "No se pudo enrutar el pedido a producción.", result.CodigoError ?? "PEDIDO_ROUTING_COCINA_ERROR");
            }
        }

        return Result.Ok(pedido);
    }

    private static string GenerarCodigoTemporal()
        => $"TMP{DateTime.UtcNow:HHmmss}";

    private static string SiguienteItem(int index)
        => (index + 1).ToString("000");
}
