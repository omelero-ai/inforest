using Inforest.Application.Interfaces;
using Inforest.Domain.Common;

namespace Inforest.Application.Ventas;

// ── Commands ────────────────────────────────────────────────────────────────

/// <summary>
/// Modifica la cortesía de un pedido/documento.
/// Legacy: FrmActualizarPedidos.BtnActualizar1_Click — usp_ActualizarCabPeDoc @opcion='1'. BR-ACTPED-001/002.
/// </summary>
public sealed record ActualizarCortesiaPedidoCommand(
    string CodigoDocumento,
    string CodigoCortesiaAnterior,
    string CodigoCortesiaNueva,
    string CodigoUsuario);

/// <summary>
/// Modifica el canal de venta de un pedido/documento.
/// Legacy: FrmActualizarPedidos.Command1_Click — usp_ActualizarCabPeDoc @opcion='2'. BR-ACTPED-001.
/// </summary>
public sealed record ActualizarCanalVentaPedidoCommand(
    string CodigoDocumento,
    string CodigoCanalAnterior,
    string CodigoCanalNuevo,
    /// <summary>Código de cliente delivery (solo cuando canal='02' — Delivery).</summary>
    string CodigoClienteDelivery,
    string CodigoUsuario);

// ── Query ────────────────────────────────────────────────────────────────────

/// <summary>Obtiene los datos actuales del pedido/documento para mostrar en el formulario.</summary>
public sealed record ObtenerDatosPedidoActualizarQuery(string CodigoDocumento);

/// <summary>Resultado de la consulta inicial del formulario.</summary>
public sealed record DatosPedidoActualizar(
    string CodigoPedido,
    string CodigoCortesia,
    string DescripcionCortesia,
    string CodigoCanalVenta,
    string DescripcionCanalVenta);

/// <summary>Opción de catálogo (cortesía o canal de venta).</summary>
public sealed record OpcionCatalogo(string Codigo, string Descripcion);

/// <summary>Obtiene catálogos para los combos de cortesías y canales de venta.</summary>
public sealed record ObtenerCatalogosCambioQuery();

public sealed record CatalogosCambioResultado(
    IReadOnlyList<OpcionCatalogo> Cortesias,
    IReadOnlyList<OpcionCatalogo> CanalesVenta);

// ── Repository Interface ──────────────────────────────────────────────────────

/// <summary>Contrato de acceso a datos para actualizar datos del pedido.</summary>
public interface IActualizarDatosPedidoRepository
{
    Task<DatosPedidoActualizar?> ObtenerDatosActualesAsync(string codigoDocumento, CancellationToken ct = default);
    Task<IReadOnlyList<OpcionCatalogo>> ObtenerCortesiasAsync(CancellationToken ct = default);
    Task<IReadOnlyList<OpcionCatalogo>> ObtenerCanalesVentaAsync(CancellationToken ct = default);

    /// <summary>
    /// Actualiza la cortesía: usp_ActualizarCabPeDoc @opcion='1'. BR-ACTPED-001.
    /// </summary>
    Task ActualizarCortesiaAsync(
        string codigoDocumento,
        string codigoCortesiaNueva,
        string codigoCortesiaAnterior,
        string codigoUsuario,
        CancellationToken ct = default);

    /// <summary>
    /// Actualiza el canal de venta: usp_ActualizarCabPeDoc @opcion='2'. BR-ACTPED-001.
    /// </summary>
    Task ActualizarCanalVentaAsync(
        string codigoDocumento,
        string codigoCanalNuevo,
        string codigoClienteDelivery,
        string codigoCanalAnterior,
        string codigoUsuario,
        CancellationToken ct = default);
}

// ── Handlers ─────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene los datos actuales de un pedido/documento y los catálogos para el formulario.
/// Legacy: FrmActualizarPedidos.iniciar()
/// </summary>
public sealed class ObtenerDatosPedidoActualizarHandler
{
    private readonly IActualizarDatosPedidoRepository _repo;

    public ObtenerDatosPedidoActualizarHandler(IActualizarDatosPedidoRepository repo)
        => _repo = repo;

    public async Task<Result<(DatosPedidoActualizar Datos, IReadOnlyList<OpcionCatalogo> Cortesias, IReadOnlyList<OpcionCatalogo> Canales)>>
        HandleAsync(ObtenerDatosPedidoActualizarQuery query, CancellationToken ct = default)
    {
        try
        {
            var datos = await _repo.ObtenerDatosActualesAsync(query.CodigoDocumento, ct);
            if (datos is null)
                return Result.Fail<(DatosPedidoActualizar, IReadOnlyList<OpcionCatalogo>, IReadOnlyList<OpcionCatalogo>)>(
                    "No se encontró el documento.", "ACTPED_DOC_NO_ENCONTRADO");

            var cortesias = await _repo.ObtenerCortesiasAsync(ct);
            var canales = await _repo.ObtenerCanalesVentaAsync(ct);
            return Result.Ok((datos, cortesias, canales));
        }
        catch (Exception ex)
        {
            return Result.Fail<(DatosPedidoActualizar, IReadOnlyList<OpcionCatalogo>, IReadOnlyList<OpcionCatalogo>)>(
                ex.Message, "ACTPED_ERROR");
        }
    }
}

/// <summary>
/// Modifica la cortesía del pedido. Requiere supervisor nivel 31 (validado en UI).
/// Legacy: FrmActualizarPedidos.BtnActualizar1_Click.
/// <para>BR-ACTPED-001: Requiere supervisor (nivel 31).</para>
/// <para>BR-ACTPED-002: No se puede modificar cortesía si el documento no tiene cortesía asignada.</para>
/// </summary>
public sealed class ActualizarCortesiaPedidoHandler
{
    private readonly IActualizarDatosPedidoRepository _repo;

    public ActualizarCortesiaPedidoHandler(IActualizarDatosPedidoRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarCortesiaPedidoCommand command, CancellationToken ct = default)
    {
        // BR-ACTPED-002: No se puede cambiar cortesía si no tiene cortesía previa
        if (string.IsNullOrWhiteSpace(command.CodigoCortesiaAnterior))
            return Result.Fail(
                "El documento no tiene cortesía asignada; no se puede modificar.",
                "ACTPED_SIN_CORTESIA");

        if (string.IsNullOrWhiteSpace(command.CodigoCortesiaNueva))
            return Result.Fail("Debe seleccionar una cortesía de reemplazo.", "ACTPED_CORTESIA_REQUERIDA");

        try
        {
            await _repo.ActualizarCortesiaAsync(
                command.CodigoDocumento,
                command.CodigoCortesiaNueva,
                command.CodigoCortesiaAnterior,
                command.CodigoUsuario,
                ct);
            return Result.Ok();
        }
        catch (Exception ex)
        {
            return Result.Fail(ex.Message, "ACTPED_CORTESIA_ERROR");
        }
    }
}

/// <summary>
/// Modifica el canal de venta del pedido. Requiere supervisor nivel 32 (validado en UI).
/// Legacy: FrmActualizarPedidos.Command1_Click.
/// <para>BR-ACTPED-001: Requiere supervisor (nivel 32).</para>
/// <para>BR-ACTPED-003: Si el canal es '02' (Delivery), se requiere cliente delivery.</para>
/// </summary>
public sealed class ActualizarCanalVentaPedidoHandler
{
    private readonly IActualizarDatosPedidoRepository _repo;

    public ActualizarCanalVentaPedidoHandler(IActualizarDatosPedidoRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarCanalVentaPedidoCommand command, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(command.CodigoCanalNuevo))
            return Result.Fail("Debe seleccionar un canal de venta.", "ACTPED_CANAL_REQUERIDO");

        // BR-ACTPED-003: Canal delivery requiere cliente
        if (command.CodigoCanalNuevo == "02" && string.IsNullOrWhiteSpace(command.CodigoClienteDelivery))
            return Result.Fail(
                "El canal Delivery requiere un cliente delivery asignado.",
                "ACTPED_CANAL_DELIVERY_SIN_CLIENTE");

        try
        {
            await _repo.ActualizarCanalVentaAsync(
                command.CodigoDocumento,
                command.CodigoCanalNuevo,
                command.CodigoClienteDelivery,
                command.CodigoCanalAnterior,
                command.CodigoUsuario,
                ct);
            return Result.Ok();
        }
        catch (Exception ex)
        {
            return Result.Fail(ex.Message, "ACTPED_CANAL_ERROR");
        }
    }
}
