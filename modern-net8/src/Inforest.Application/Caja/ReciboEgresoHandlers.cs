using Inforest.Domain.Common;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Caja;

// ── Queries ───────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene recibos de egreso en un rango de fechas.
/// Legacy: frmReciboEgreso.frm Form_Load / cmdProcesa_Click. BR-RECIBO-006.
/// </summary>
public sealed record ObtenerEgresosQuery(
    DateTime Desde,
    DateTime Hasta,
    string? UsuarioFiltro = null);

public sealed class ObtenerEgresosHandler
{
    private readonly IReciboEgresoRepository _repository;

    public ObtenerEgresosHandler(IReciboEgresoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<ReciboEgreso>>> HandleAsync(
        ObtenerEgresosQuery query, CancellationToken ct = default)
    {
        if (query.Desde > query.Hasta)
            return Result.Fail<IReadOnlyList<ReciboEgreso>>(
                "La fecha de inicio no puede ser posterior a la fecha fin.", "EGRESO_RANGO_FECHAS_INVALIDO");

        var items = await _repository.ObtenerPorFechaAsync(
            query.Desde, query.Hasta, query.UsuarioFiltro, ct);

        return Result.Ok<IReadOnlyList<ReciboEgreso>>(items);
    }
}

// ── Commands ──────────────────────────────────────────────────────────────────

/// <summary>
/// Registra un nuevo recibo de egreso de caja.
/// Legacy: frmReciboEgresoDetalle.frm Case Is=1 (Grabar). BR-RECIBO-007, BR-RECIBO-010, BR-RECIBO-011, BR-RECIBO-012.
/// </summary>
public sealed record RegistrarEgresoCommand(
    DateTime Fecha,
    decimal Monto,
    string Descripcion,
    string ModoEgreso = "VARIOS",
    string Moneda = "01",
    decimal TipoCambio = 1m,
    string? Autorizacion = null,
    string? TipoEgreso = null,
    string? Referencia = null,
    string? Turno = null,
    string? Caja = null,
    string? Usuario = null,
    DateTime? FechaDiaContable = null,
    string? TipoIdentificacion = null,
    string? Identificacion = null,
    string? RazonSocial = null,
    string? NumeroComprobante = null,
    string? TipoComprobante = null,
    string? CentroCosto = null,
    string? Observacion = null);

public sealed class RegistrarEgresoHandler
{
    private readonly IReciboEgresoRepository _repository;

    public RegistrarEgresoHandler(IReciboEgresoRepository repository)
        => _repository = repository;

    public async Task<Result<ReciboEgreso>> HandleAsync(
        RegistrarEgresoCommand command, CancellationToken ct = default)
    {
        // BR-RECIBO-007: Calcular correlativo YY + 8 dígitos
        var anoCorto = DateTime.Now.ToString("yy");
        var maxCorrelativo = await _repository.ObtenerMaximoCorrelativoAnioAsync(anoCorto, ct);

        string nuevoCorrelativo;
        if (string.IsNullOrEmpty(maxCorrelativo) || !maxCorrelativo.StartsWith(anoCorto))
        {
            nuevoCorrelativo = anoCorto + "00000001";
        }
        else
        {
            var secuencia = maxCorrelativo[2..];
            if (long.TryParse(secuencia, out var num))
                nuevoCorrelativo = anoCorto + (num + 1).ToString().PadLeft(8, '0');
            else
                nuevoCorrelativo = anoCorto + "00000001";
        }

        ReciboEgreso recibo;
        try
        {
            recibo = ReciboEgreso.Registrar(
                nuevoCorrelativo,
                command.Fecha,
                command.Monto,
                command.Descripcion,
                command.ModoEgreso,
                command.Moneda,
                command.TipoCambio,
                command.Autorizacion,
                command.TipoEgreso,
                command.Referencia,
                command.Turno,
                command.Caja,
                command.Usuario,
                command.FechaDiaContable,
                command.TipoIdentificacion,
                command.Identificacion,
                command.RazonSocial,
                command.NumeroComprobante,
                command.TipoComprobante,
                command.CentroCosto,
                command.Observacion);
        }
        catch (DomainException ex)
        {
            return Result.Fail<ReciboEgreso>(ex.Message, ex.Code);
        }

        await _repository.InsertarAsync(recibo, ct);
        return Result.Ok(recibo);
    }
}

/// <summary>
/// Anula un recibo de egreso existente.
/// Legacy: frmReciboEgresoDetalle.frm Case Is=2 (Anular). BR-RECIBO-008.
/// BR-RECIBO-009: La autorización (lSolicitaClaveEgreso) se verifica en la capa de presentación.
/// </summary>
public sealed record AnularEgresoCommand(string CodigoRecibo);

public sealed class AnularEgresoHandler
{
    private readonly IReciboEgresoRepository _repository;

    public AnularEgresoHandler(IReciboEgresoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(AnularEgresoCommand command, CancellationToken ct = default)
    {
        var recibo = await _repository.ObtenerPorCodigoAsync(command.CodigoRecibo, ct);
        if (recibo is null)
            return Result.Fail("Recibo de egreso no encontrado.", "RECIBO_EGRESO_NO_ENCONTRADO");

        try
        {
            recibo.Anular();
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }

        await _repository.ActualizarEstadoAsync(command.CodigoRecibo, "04", ct);
        return Result.Ok();
    }
}
