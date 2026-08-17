using Inforest.Application.Configuracion;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Despacho;

/// <summary>
/// Query para obtener pedidos del panel despachador.
/// <para>
/// Legacy: <c>frmDespachador.frm → cmdProcesa_Click</c>.
/// SQL: <c>select * from vDespachador ... order by lEmpacador, fFecha asc</c>.
/// </para>
/// </summary>
public sealed record ObtenerPedidosDespachadorQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string? CodigoCaja = null);

/// <summary>Handler de <see cref="ObtenerPedidosDespachadorQuery"/>.</summary>
public sealed class ObtenerPedidosDespachadorHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerPedidosDespachadorHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDespachadorResumen>>> HandleAsync(
        ObtenerPedidosDespachadorQuery query, CancellationToken ct = default)
    {
        var rows = await _repo.ObtenerResumenDespachadorAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(rows.AsEnumerable());
    }
}

/// <summary>
/// Query para exportar datos del despachador (HTML).
/// <para>
/// Legacy: <c>frmDespachador.frm → cmdExporta_Click</c>.
/// </para>
/// </summary>
public sealed record ExportarDespachadorQuery(
    DateTime FechaInicio,
    DateTime FechaFin,
    string? CodigoCaja = null);

/// <summary>Handler de <see cref="ExportarDespachadorQuery"/>.</summary>
public sealed class ExportarDespachadorHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ExportarDespachadorHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDespachadorResumen>>> HandleAsync(
        ExportarDespachadorQuery query, CancellationToken ct = default)
    {
        var rows = await _repo.ObtenerResumenDespachadorAsync(query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(rows.AsEnumerable());
    }
}

/// <summary>
/// Query de catálogo de motorizados activos para el panel de despacho.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(0)</c> con <c>vMotorizado</c>.
/// </summary>
public sealed record ObtenerMotorizadosDespachoQuery;

/// <summary>Handler de <see cref="ObtenerMotorizadosDespachoQuery"/>.</summary>
public sealed class ObtenerMotorizadosDespachoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerMotorizadosDespachoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<OperadorDespachoItem>>> HandleAsync(
        ObtenerMotorizadosDespachoQuery _,
        CancellationToken ct = default)
    {
        var items = await _repo.ObtenerMotorizadosActivosDespachoAsync(ct);
        return Result.Ok(items);
    }
}

/// <summary>
/// Query de catálogo de empacadores activos para el panel de despacho.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(8)</c> con <c>vEmpacador</c>.
/// </summary>
public sealed record ObtenerEmpacadoresDespachoQuery;

/// <summary>Handler de <see cref="ObtenerEmpacadoresDespachoQuery"/>.</summary>
public sealed class ObtenerEmpacadoresDespachoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerEmpacadoresDespachoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<OperadorDespachoItem>>> HandleAsync(
        ObtenerEmpacadoresDespachoQuery _,
        CancellationToken ct = default)
    {
        var items = await _repo.ObtenerEmpacadoresActivosDespachoAsync(ct);
        return Result.Ok(items);
    }
}

/// <summary>
/// Asigna motorizado a un pedido del panel de despacho.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(0)</c>.
/// Reglas: BR-DEL-DESP-001, BR-DEL-DESP-002, BR-DEL-DESP-003.
/// </summary>
public sealed record AsignarMotorizadoDespachoCommand(
    string CodigoPedido,
    string CodigoMotorizado,
    string Usuario,
    bool AutorizarTarifaExtra = false);

/// <summary>Handler de <see cref="AsignarMotorizadoDespachoCommand"/>.</summary>
public sealed class AsignarMotorizadoDespachoHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;
    private readonly IMotorizadoRepository _motorizadoRepo;
    private readonly IParametroRepository _parametroRepo;

    public AsignarMotorizadoDespachoHandler(
        IPedidoDeliveryRepository pedidoRepo,
        IMotorizadoRepository motorizadoRepo,
        IParametroRepository parametroRepo)
    {
        _pedidoRepo = pedidoRepo;
        _motorizadoRepo = motorizadoRepo;
        _parametroRepo = parametroRepo;
    }

    public async Task<Result> HandleAsync(AsignarMotorizadoDespachoCommand cmd, CancellationToken ct = default)
    {
        var fila = await _pedidoRepo.ObtenerResumenDespachadorPorPedidoAsync(cmd.CodigoPedido, ct);
        if (fila is null)
            return Result.Fail("No se encontró el pedido seleccionado.", "DESPACHADOR_PEDIDO_NO_ENCONTRADO");

        if (string.IsNullOrWhiteSpace(fila.Empacador))
            return Result.Fail("Debe asignar primero el empacador.", "DESPACHADOR_REQUIERE_EMPACADOR");

        if (fila.FechaAsignacion.HasValue || (!string.IsNullOrWhiteSpace(fila.Motorizado) && fila.Motorizado != "0000"))
            return Result.Fail("El pedido ya fue asignado.", "DESPACHADOR_PEDIDO_YA_ASIGNADO");

        var motorizado = await _motorizadoRepo.ObtenerPorCodigoAsync(cmd.CodigoMotorizado, ct);
        if (motorizado is null)
            return Result.Fail("Motorizado no encontrado.", "MOTORIZADO_NO_ENCONTRADO");

        var configuracion = await _parametroRepo.ObtenerConfiguracionAsync(ct);
        var tarifaDiaria = configuracion?.tTarifaActualMotorizado switch
        {
            "Tarifa Dom" => motorizado.TarifaSabadoDomingo,
            "Tarifa Especial" => motorizado.TarifaEspecial,
            _ => motorizado.TarifaLunesViernes
        };

        var asignaciones = await _pedidoRepo.ContarAsignacionesPrincipalesMotorizadoAsync(cmd.CodigoMotorizado, DateTime.Today, ct);
        var reasignaciones = await _pedidoRepo.ContarReasignacionesMotorizadoAsync(cmd.CodigoMotorizado, DateTime.Today, ct);
        var montoAsignado = decimal.Round((tarifaDiaria * (asignaciones + reasignaciones)) + tarifaDiaria, 2, MidpointRounding.AwayFromZero);

        var superaMontoMaximo = configuracion is not null &&
                                configuracion.nAsignacionMotorizado > 0 &&
                                montoAsignado > Convert.ToDecimal(configuracion.nAsignacionMotorizado);

        if (superaMontoMaximo && !cmd.AutorizarTarifaExtra)
            return Result.Fail(
                "Con la asignación de este pedido se supera el monto máximo de asignación por motorizado.",
                "DESPACHADOR_SUPERA_MONTO_MAXIMO");

        if (configuracion is not null &&
            int.TryParse(configuracion.tMaxMotorizado, out var maxMotorizado) &&
            maxMotorizado > 0)
        {
            var activos = await _pedidoRepo.ContarPedidosActivosMotorizadoAsync(cmd.CodigoMotorizado, ct);
            if (maxMotorizado < activos + 1)
            {
                return Result.Fail(
                    $"Ha alcanzado el tope máximo de asignación por motorizado: {maxMotorizado}.",
                    "DESPACHADOR_TOPE_MOTORIZADO");
            }
        }

        await _pedidoRepo.AsignarMotorizadoDespachoAsync(
            cmd.CodigoPedido,
            cmd.CodigoMotorizado,
            tarifaDiaria,
            superaMontoMaximo,
            ct);

        return Result.Ok();
    }
}

/// <summary>
/// Desasigna motorizado desde el panel de despacho.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(1)</c>.
/// </summary>
public sealed record DesasignarMotorizadoDespachoCommand(string CodigoPedido);

/// <summary>Handler de <see cref="DesasignarMotorizadoDespachoCommand"/>.</summary>
public sealed class DesasignarMotorizadoDespachoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public DesasignarMotorizadoDespachoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(DesasignarMotorizadoDespachoCommand cmd, CancellationToken ct = default)
    {
        var fila = await _repo.ObtenerResumenDespachadorPorPedidoAsync(cmd.CodigoPedido, ct);
        if (fila is null)
            return Result.Fail("No se encontró el pedido seleccionado.", "DESPACHADOR_PEDIDO_NO_ENCONTRADO");

        if (!fila.FechaAsignacion.HasValue || string.IsNullOrWhiteSpace(fila.Motorizado) || fila.Motorizado == "0000")
            return Result.Fail("El pedido no está asignado.", "DESPACHADOR_PEDIDO_NO_ASIGNADO");

        var activos = await _repo.ContarPedidosActivosMotorizadoAsync(fila.Motorizado, ct);
        var liberarBandera = activos <= 1;
        await _repo.DesasignarMotorizadoDespachoAsync(cmd.CodigoPedido, fila.Motorizado, liberarBandera, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Asigna empacador a pedido.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(8)</c>.
/// </summary>
public sealed record AsignarEmpacadorDespachoCommand(string CodigoPedido, string CodigoEmpacador);

/// <summary>Handler de <see cref="AsignarEmpacadorDespachoCommand"/>.</summary>
public sealed class AsignarEmpacadorDespachoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public AsignarEmpacadorDespachoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(AsignarEmpacadorDespachoCommand cmd, CancellationToken ct = default)
    {
        var fila = await _repo.ObtenerResumenDespachadorPorPedidoAsync(cmd.CodigoPedido, ct);
        if (fila is null)
            return Result.Fail("No se encontró el pedido seleccionado.", "DESPACHADOR_PEDIDO_NO_ENCONTRADO");

        if (!string.IsNullOrWhiteSpace(fila.Empacador))
            return Result.Fail("El pedido ya fue empacado.", "DESPACHADOR_PEDIDO_YA_EMPACADO");

        await _repo.AsignarEmpacadorDespachoAsync(cmd.CodigoPedido, cmd.CodigoEmpacador, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Desasigna empacador de pedido.
/// Legacy: <c>frmDespachador.frm → cmdOpcion_Click(9)</c>.
/// </summary>
public sealed record DesasignarEmpacadorDespachoCommand(string CodigoPedido);

/// <summary>Handler de <see cref="DesasignarEmpacadorDespachoCommand"/>.</summary>
public sealed class DesasignarEmpacadorDespachoHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public DesasignarEmpacadorDespachoHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(DesasignarEmpacadorDespachoCommand cmd, CancellationToken ct = default)
    {
        var fila = await _repo.ObtenerResumenDespachadorPorPedidoAsync(cmd.CodigoPedido, ct);
        if (fila is null)
            return Result.Fail("No se encontró el pedido seleccionado.", "DESPACHADOR_PEDIDO_NO_ENCONTRADO");

        if (string.IsNullOrWhiteSpace(fila.Empacador))
            return Result.Fail("El pedido no está empaquetado.", "DESPACHADOR_PEDIDO_NO_EMPACADO");

        if (!string.IsNullOrWhiteSpace(fila.Motorizado) && fila.Motorizado != "0000")
            return Result.Fail("El pedido ya está en camino.", "DESPACHADOR_PEDIDO_EN_CAMINO");

        await _repo.DesasignarEmpacadorDespachoAsync(cmd.CodigoPedido, ct);
        return Result.Ok();
    }
}
