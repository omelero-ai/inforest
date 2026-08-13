using Inforest.Domain.Common;
using Inforest.Domain.Entities.Configuracion;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Repositories;
using Inforest.Application.Configuracion;
using MotorizadoEntity = Inforest.Domain.Entities.Motorizado.Motorizado;

namespace Inforest.Application.Motorizado;

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para obtener todos los motorizados activos.
/// <para>
/// Legacy: vista <c>vMotorizado</c> en <c>frmMotorizado.frm</c> y <c>frmLlegadaSalida.frm</c>.
/// </para>
/// </summary>
public sealed record ObtenerMotorizadosActivosQuery;

/// <summary>Handler de <see cref="ObtenerMotorizadosActivosQuery"/>.</summary>
public sealed class ObtenerMotorizadosActivosHandler
{
    private readonly IMotorizadoRepository _repo;

    public ObtenerMotorizadosActivosHandler(IMotorizadoRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<MotorizadoEntity>>> HandleAsync(
        ObtenerMotorizadosActivosQuery _, CancellationToken ct = default)
    {
        var motorizados = await _repo.ObtenerTodosAsync(ct);
        return Result.Ok(motorizados);
    }
}

/// <summary>
/// Query para obtener pedidos disponibles en la pantalla de reasignación.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm → Form_Load/cmdBuscar_Click</c>.
/// </para>
/// </summary>
public sealed record ObtenerPedidosReasignacionQuery(DateTime FechaInicio, DateTime FechaFin, string? CriterioPedido);

/// <summary>Handler de <see cref="ObtenerPedidosReasignacionQuery"/>.</summary>
public sealed class ObtenerPedidosReasignacionHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerPedidosReasignacionHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoReasignacionMotorizado>>> HandleAsync(
        ObtenerPedidosReasignacionQuery query,
        CancellationToken ct = default)
    {
        if (query.FechaInicio > query.FechaFin)
            return Result.Fail<IEnumerable<PedidoReasignacionMotorizado>>(
                "Error en rango de fechas.",
                "REASIGNACION_RANGO_FECHAS_INVALIDO");

        var pedidos = await _repo.ObtenerPedidosReasignacionAsync(
            query.FechaInicio,
            query.FechaFin,
            query.CriterioPedido,
            ct);

        return Result.Ok(pedidos);
    }
}

/// <summary>
/// Query para obtener la tarifa de un motorizado según tipo de día.
/// <para>
/// Legacy: <c>TMOTORIZADODATOS</c> consultado en <c>frmTarifaMotorizado.frm</c>.
/// </para>
/// Regla BR-DEL-002.
/// </summary>
public sealed record ObtenerTarifaMotorizadoQuery(string CodigoMotorizado, TipoDiaTarifa TipoDia);

/// <summary>Handler de <see cref="ObtenerTarifaMotorizadoQuery"/>.</summary>
public sealed class ObtenerTarifaMotorizadoHandler
{
    private readonly IMotorizadoRepository _repo;

    public ObtenerTarifaMotorizadoHandler(IMotorizadoRepository repo)
        => _repo = repo;

    public async Task<Result<decimal>> HandleAsync(ObtenerTarifaMotorizadoQuery query, CancellationToken ct = default)
    {
        var motorizado = await _repo.ObtenerPorCodigoAsync(query.CodigoMotorizado, ct);
        if (motorizado is null)
            return Result.Fail<decimal>("Motorizado no encontrado.", "MOTORIZADO_NO_ENCONTRADO");

        var tarifa = motorizado.ObtenerTarifaPorDia(query.TipoDia);
        return Result.Ok(tarifa);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para asignar un motorizado a un pedido delivery.
/// <para>
/// Legacy: <c>frmAsignacionMotorizado.frm → Form_Load</c> — asigna motorizado y cambia estado del pedido.
/// </para>
/// Regla BR-DEL-001.
/// </summary>
public sealed record AsignarMotorizadoCommand(
    string CodigoPedido,
    string CodigoMotorizado,
    string UsuarioAsignacion);

/// <summary>Handler de <see cref="AsignarMotorizadoCommand"/>.</summary>
public sealed class AsignarMotorizadoHandler
{
    private readonly IMotorizadoRepository _motorizadoRepo;
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public AsignarMotorizadoHandler(
        IMotorizadoRepository motorizadoRepo,
        IPedidoDeliveryRepository pedidoRepo)
    {
        _motorizadoRepo = motorizadoRepo;
        _pedidoRepo = pedidoRepo;
    }

    public async Task<Result> HandleAsync(AsignarMotorizadoCommand cmd, CancellationToken ct = default)
    {
        var motorizado = await _motorizadoRepo.ObtenerPorCodigoAsync(cmd.CodigoMotorizado, ct);
        if (motorizado is null)
            return Result.Fail("Motorizado no encontrado.", "MOTORIZADO_NO_ENCONTRADO");

        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("Pedido no encontrado.", "PEDIDO_NO_ENCONTRADO");

        pedido.AsignarMotorizado(cmd.CodigoMotorizado);

        var asignacion = AsignacionMotorizado.Crear(cmd.CodigoPedido, cmd.CodigoMotorizado, cmd.UsuarioAsignacion);
        await _motorizadoRepo.RegistrarAsignacionAsync(asignacion, ct);
        await _pedidoRepo.AsignarMotorizadoAsync(cmd.CodigoPedido, cmd.CodigoMotorizado, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para reasignar motorizado.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm</c>.
/// </para>
/// </summary>
public sealed record ReasignarMotorizadoCommand(
    string CodigoPedido,
    string NuevoCodigoMotorizado,
    string UsuarioReasignacion);

/// <summary>Handler de <see cref="ReasignarMotorizadoCommand"/>.</summary>
public sealed class ReasignarMotorizadoHandler
{
    private readonly IMotorizadoRepository _repo;
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public ReasignarMotorizadoHandler(IMotorizadoRepository repo, IPedidoDeliveryRepository pedidoRepo)
    {
        _repo = repo;
        _pedidoRepo = pedidoRepo;
    }

    public async Task<Result> HandleAsync(ReasignarMotorizadoCommand cmd, CancellationToken ct = default)
    {
        var asignacionActual = await _repo.ObtenerAsignacionActivaAsync(cmd.CodigoPedido, ct);
        if (asignacionActual is null)
            return Result.Fail("No existe asignación activa para el pedido.", "REASIGNACION_SIN_ASIGNACION");

        var nuevaAsignacion = asignacionActual.Reasignar(cmd.NuevoCodigoMotorizado, cmd.UsuarioReasignacion);
        await _repo.RegistrarAsignacionAsync(nuevaAsignacion, ct);
        await _pedidoRepo.AsignarMotorizadoAsync(cmd.CodigoPedido, cmd.NuevoCodigoMotorizado, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para asignar motorizado en el flujo de reasignación delivery.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm → cmdOpcion_Click(Case 1)</c>.
/// Actualiza <c>tMotorizadoN</c>, <c>nTarifaMotorizadoN</c> y
/// <c>nTarifaExtraN</c> respetando <c>TPARAMETRO.nAsignacionMotorizado</c>.
/// </para>
/// </summary>
public sealed record AsignarReasignacionMotorizadoCommand(
    string CodigoPedido,
    string CodigoMotorizado,
    string UsuarioAsignacion,
    bool AutorizarTarifaExtra = false);

/// <summary>Handler de <see cref="AsignarReasignacionMotorizadoCommand"/>.</summary>
public sealed class AsignarReasignacionMotorizadoHandler
{
    private readonly IMotorizadoRepository _motorizadoRepo;
    private readonly IPedidoDeliveryRepository _pedidoRepo;
    private readonly IParametroRepository _parametroRepository;

    public AsignarReasignacionMotorizadoHandler(
        IMotorizadoRepository motorizadoRepo,
        IPedidoDeliveryRepository pedidoRepo,
        IParametroRepository parametroRepository)
    {
        _motorizadoRepo = motorizadoRepo;
        _pedidoRepo = pedidoRepo;
        _parametroRepository = parametroRepository;
    }

    public async Task<Result> HandleAsync(AsignarReasignacionMotorizadoCommand cmd, CancellationToken ct = default)
    {
        var motorizado = await _motorizadoRepo.ObtenerPorCodigoAsync(cmd.CodigoMotorizado, ct);
        if (motorizado is null)
            return Result.Fail("Motorizado no encontrado.", "MOTORIZADO_NO_ENCONTRADO");

        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("Pedido no encontrado.", "PEDIDO_NO_ENCONTRADO");

        var configuracion = await _parametroRepository.ObtenerConfiguracionAsync(ct);
        var tarifaDiaria = ObtenerTarifaReasignacion(configuracion, motorizado);

        var asignacionesPrincipales = await _pedidoRepo.ContarAsignacionesPrincipalesMotorizadoAsync(cmd.CodigoMotorizado, DateTime.Today, ct);
        var reasignaciones = await _pedidoRepo.ContarReasignacionesMotorizadoAsync(cmd.CodigoMotorizado, DateTime.Today, ct);
        var montoAsignado = decimal.Round((tarifaDiaria * (asignacionesPrincipales + reasignaciones)) + tarifaDiaria, 2, MidpointRounding.AwayFromZero);

        var superaMaximo = configuracion is not null &&
                           configuracion.nAsignacionMotorizado > 0 &&
                           montoAsignado > Convert.ToDecimal(configuracion.nAsignacionMotorizado);

        if (superaMaximo && !cmd.AutorizarTarifaExtra)
            return Result.Fail(
                "La asignación supera el monto máximo permitido por motorizado.",
                "REASIGNACION_SUPERA_MONTO_MAXIMO");

        await _pedidoRepo.ActualizarReasignacionMotorizadoAsync(
            cmd.CodigoPedido,
            cmd.CodigoMotorizado,
            tarifaDiaria,
            superaMaximo,
            ct);

        return Result.Ok();
    }

    private static decimal ObtenerTarifaReasignacion(ConfiguracionSistema? configuracion, MotorizadoEntity motorizado)
        => configuracion?.tTarifaActualMotorizado switch
        {
            "Tarifa Dom" => motorizado.TarifaSabadoDomingo,
            "Tarifa Especial" => motorizado.TarifaEspecial,
            _ => motorizado.TarifaLunesViernes
        };
}

/// <summary>
/// Comando para limpiar la reasignación de motorizado de un pedido.
/// <para>
/// Legacy: <c>frmReasignacionMotorizado.frm → cmdOpcion_Click(Case 2)</c>.
/// </para>
/// </summary>
public sealed record DesasignarReasignacionMotorizadoCommand(string CodigoPedido);

/// <summary>Handler de <see cref="DesasignarReasignacionMotorizadoCommand"/>.</summary>
public sealed class DesasignarReasignacionMotorizadoHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public DesasignarReasignacionMotorizadoHandler(IPedidoDeliveryRepository pedidoRepo)
        => _pedidoRepo = pedidoRepo;

    public async Task<Result> HandleAsync(DesasignarReasignacionMotorizadoCommand cmd, CancellationToken ct = default)
    {
        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("Pedido no encontrado.", "PEDIDO_NO_ENCONTRADO");

        await _pedidoRepo.LimpiarReasignacionMotorizadoAsync(cmd.CodigoPedido, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para registrar llegada/salida de motorizado.
/// <para>
/// Legacy: <c>frmLlegadaSalida.frm</c> — pantalla de control de salida/llegada.
/// Es el formulario de startup del exe Motorizado.
/// </para>
/// Regla BR-DEL-006.
/// </summary>
public sealed record RegistrarLlegadaSalidaCommand(
    string CodigoPedido,
    string CodigoMotorizado,
    bool EsLlegada);

/// <summary>Handler de <see cref="RegistrarLlegadaSalidaCommand"/>.</summary>
public sealed class RegistrarLlegadaSalidaHandler
{
    private readonly IMotorizadoRepository _repo;

    public RegistrarLlegadaSalidaHandler(IMotorizadoRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(RegistrarLlegadaSalidaCommand cmd, CancellationToken ct = default)
    {
        if (cmd.EsLlegada)
            await _repo.RegistrarLlegadaAsync(cmd.CodigoPedido, cmd.CodigoMotorizado, ct);
        else
            await _repo.RegistrarSalidaAsync(cmd.CodigoPedido, cmd.CodigoMotorizado, ct);

        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar la tarifa activa del motorizado.
/// <para>
/// Legacy: <c>frmTarifaMotorizado.frm → cmdOpcion_Click(Case 1)</c> —
/// actualiza <c>TPARAMETRO.tTarifaActualMotorizado</c> con supervisión.
/// Valida <c>Supervisor("21")</c> antes de grabar.
/// </para>
/// Regla BR-DEL-005.
/// </summary>
public sealed record ActualizarTarifaMotorizadoCommand(
    string CodigoMotorizado,
    decimal TarifaLunesViernes,
    decimal TarifaSabadoDomingo,
    decimal TarifaEspecial,
    string UsuarioModificacion);

/// <summary>Handler de <see cref="ActualizarTarifaMotorizadoCommand"/>.</summary>
public sealed class ActualizarTarifaMotorizadoHandler
{
    private readonly IMotorizadoRepository _repo;

    public ActualizarTarifaMotorizadoHandler(IMotorizadoRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarTarifaMotorizadoCommand cmd, CancellationToken ct = default)
    {
        var motorizado = await _repo.ObtenerPorCodigoAsync(cmd.CodigoMotorizado, ct);
        if (motorizado is null)
            return Result.Fail("Motorizado no encontrado.", "MOTORIZADO_NO_ENCONTRADO");

        motorizado.ActualizarTarifas(cmd.TarifaLunesViernes, cmd.TarifaSabadoDomingo, cmd.TarifaEspecial);
        await _repo.ActualizarTarifasAsync(
            cmd.CodigoMotorizado,
            cmd.TarifaLunesViernes,
            cmd.TarifaSabadoDomingo,
            cmd.TarifaEspecial,
            cmd.UsuarioModificacion,
            ct);
        return Result.Ok();
    }
}
