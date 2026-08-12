using Inforest.Domain.Common;
using Inforest.Application.Interfaces;

namespace Inforest.Application.Delivery;

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para obtener órdenes externas pendientes (Rappi u otro canal).
/// <para>
/// Legacy: <c>frmOrdenesConsola.frm</c> — polling de órdenes externas.
/// Habilitado por flag <c>lOrdenesRappi</c> en <c>TCAJA</c>.
/// </para>
/// Regla BR-DEL-004.
/// </summary>
public sealed record ObtenerOrdenesExternasQuery;

/// <summary>Handler de <see cref="ObtenerOrdenesExternasQuery"/>.</summary>
public sealed class ObtenerOrdenesExternasHandler
{
    private readonly IRappiOrderService _rappiService;

    public ObtenerOrdenesExternasHandler(IRappiOrderService rappiService)
        => _rappiService = rappiService;

    public async Task<Result<IEnumerable<OrdenExterna>>> HandleAsync(
        ObtenerOrdenesExternasQuery _, CancellationToken ct = default)
    {
        if (!_rappiService.EstaHabilitado)
            return Result.Ok<IEnumerable<OrdenExterna>>([]);

        var ordenes = await _rappiService.ObtenerOrdenesPendientesAsync(ct);
        return Result.Ok(ordenes);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para enviar datos de un pedido a Rappi.
/// <para>
/// Legacy: <c>frmEnvioDatosRappi.frm</c> — serialización y envío al canal Rappi.
/// </para>
/// Regla BR-DEL-004.
/// </summary>
public sealed record EnviarDatosRappiCommand(string CodigoPedido);

/// <summary>Handler de <see cref="EnviarDatosRappiCommand"/>.</summary>
public sealed class EnviarDatosRappiHandler
{
    private readonly IRappiOrderService _rappiService;

    public EnviarDatosRappiHandler(IRappiOrderService rappiService)
        => _rappiService = rappiService;

    public async Task<Result> HandleAsync(EnviarDatosRappiCommand cmd, CancellationToken ct = default)
    {
        if (!_rappiService.EstaHabilitado)
            return Result.Fail("La integración Rappi no está habilitada para esta caja.", "RAPPI_NO_HABILITADO");

        var ok = await _rappiService.EnviarOrdenAsync(cmd.CodigoPedido, ct);
        return ok ? Result.Ok() : Result.Fail("Error al enviar la orden a Rappi.", "RAPPI_ENVIO_FALLIDO");
    }
}
