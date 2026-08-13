using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Application.Ventas;

// ─────────────────────────────────────────────────────────────────────────────
// Queries
// ─────────────────────────────────────────────────────────────────────────────

public sealed record ObtenerCorrelativosPorCajaQuery(string CodigoCaja);

/// <summary>
/// Devuelve la lista de correlativos de una caja.
/// Legacy: frmDocumentoCorrelativo.frm Form_Load — SELECT * FROM TTIPODOCUMENTOIMPRESORA WHERE tCaja=sCaja
/// BR-POS-008.
/// </summary>
public sealed class ObtenerCorrelativosPorCajaHandler
{
    private readonly ICorrelativoRepository _repository;

    public ObtenerCorrelativosPorCajaHandler(ICorrelativoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<CorrelativoDocumento>>> HandleAsync(
        ObtenerCorrelativosPorCajaQuery query,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<IReadOnlyList<CorrelativoDocumento>>(
                "El código de caja es requerido.", "CORRELATIVO_CAJA_REQUERIDA");

        var correlativos = await _repository.ObtenerPorCajaAsync(query.CodigoCaja, ct);
        return Result.Ok(correlativos);
    }
}

public sealed record ObtenerCorrelativoFacturaQuery(string CodigoCaja, string TipoEmision);

/// <summary>
/// Obtiene el correlativo de factura para el inicio del POS.
/// Legacy: modPuntoVenta.bas Sub Main — NFactura = tUltimoNumero o "Sin Correlativo". BR-POS-008.
/// </summary>
public sealed class ObtenerCorrelativoFacturaHandler
{
    private readonly ICorrelativoRepository _repository;

    public ObtenerCorrelativoFacturaHandler(ICorrelativoRepository repository)
        => _repository = repository;

    public async Task<Result<string>> HandleAsync(
        ObtenerCorrelativoFacturaQuery query,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(query.CodigoCaja))
            return Result.Fail<string>("El código de caja es requerido.", "CORRELATIVO_CAJA_REQUERIDA");

        var correlativo = await _repository.ObtenerPorCajaYTipoAsync(query.CodigoCaja, query.TipoEmision, ct);
        var numero = correlativo?.ObtenerNumeroActual() ?? "Sin Correlativo";
        return Result.Ok(numero);
    }
}

public sealed record ObtenerTodosCorrelativosQuery;

/// <summary>
/// Devuelve todos los correlativos del sistema para la pantalla de administración.
/// Legacy: frmDocumentoCorrelativo.frm — carga completa de TTIPODOCUMENTOIMPRESORA.
/// </summary>
public sealed class ObtenerTodosCorrelativosHandler
{
    private readonly ICorrelativoRepository _repository;

    public ObtenerTodosCorrelativosHandler(ICorrelativoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<CorrelativoDocumento>>> HandleAsync(
        ObtenerTodosCorrelativosQuery _,
        CancellationToken ct = default)
    {
        var correlativos = await _repository.ObtenerTodosAsync(ct);
        return Result.Ok(correlativos);
    }
}
