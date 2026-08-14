using Inforest.Domain.Common;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Caja;

// ── Queries ───────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene recibos de ingreso en un rango de fechas.
/// Legacy: frmReciboIngreso.frm Form_Load / cmdProcesa_Click. BR-RECIBO-006.
/// </summary>
public sealed record ObtenerIngresosQuery(
    DateTime Desde,
    DateTime Hasta,
    /// <summary>Código de usuario para filtro lMCPV. BR-RECIBO-006.</summary>
    string? UsuarioFiltro = null);

public sealed class ObtenerIngresosHandler
{
    private readonly IReciboIngresoRepository _repository;

    public ObtenerIngresosHandler(IReciboIngresoRepository repository)
        => _repository = repository;

    public async Task<Result<IReadOnlyList<ReciboIngreso>>> HandleAsync(
        ObtenerIngresosQuery query, CancellationToken ct = default)
    {
        if (query.Desde > query.Hasta)
            return Result.Fail<IReadOnlyList<ReciboIngreso>>(
                "La fecha de inicio no puede ser posterior a la fecha fin.", "RECIBO_RANGO_FECHAS_INVALIDO");

        var items = await _repository.ObtenerPorFechaAsync(
            query.Desde, query.Hasta, query.UsuarioFiltro, ct);

        return Result.Ok<IReadOnlyList<ReciboIngreso>>(items);
    }
}

// ── Commands ──────────────────────────────────────────────────────────────────

/// <summary>
/// Registra un nuevo recibo de ingreso de caja.
/// Legacy: frmReciboIngresoDetalle.frm Case Is=1 (Grabar). BR-RECIBO-001, BR-RECIBO-011.
/// </summary>
public sealed record RegistrarIngresoCommand(
    DateTime Fecha,
    decimal Monto,
    string Descripcion,
    string Moneda = "01",
    string TipoPago = "01",
    string? CodigoTarjeta = null,
    string? Referencia = null,
    decimal TipoCambio = 1m,
    string? Autorizacion = null,
    bool EsAnticipo = true,
    string? Turno = null,
    string? Caja = null,
    string? Usuario = null,
    DateTime? FechaDiaContable = null,
    string? CodigoPedido = null);

public sealed class RegistrarIngresoHandler
{
    private readonly IReciboIngresoRepository _repository;

    public RegistrarIngresoHandler(IReciboIngresoRepository repository)
        => _repository = repository;

    public async Task<Result<ReciboIngreso>> HandleAsync(
        RegistrarIngresoCommand command, CancellationToken ct = default)
    {
        // BR-RECIBO-001: Calcular correlativo YY + 8 dígitos
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

        ReciboIngreso recibo;
        try
        {
            recibo = ReciboIngreso.Registrar(
                nuevoCorrelativo,
                command.Fecha,
                command.Monto,
                command.Descripcion,
                command.Moneda,
                command.TipoPago,
                command.CodigoTarjeta,
                command.Referencia,
                command.TipoCambio,
                command.Autorizacion,
                command.EsAnticipo,
                command.Turno,
                command.Caja,
                command.Usuario,
                command.FechaDiaContable,
                command.CodigoPedido);
        }
        catch (DomainException ex)
        {
            return Result.Fail<ReciboIngreso>(ex.Message, ex.Code);
        }

        await _repository.InsertarAsync(recibo, ct);
        return Result.Ok(recibo);
    }
}

/// <summary>
/// Anula un recibo de ingreso existente.
/// Legacy: frmReciboIngresoDetalle.frm Case Is=2 (Anular). BR-RECIBO-003.
/// </summary>
public sealed record AnularIngresoCommand(string CodigoRecibo);

public sealed class AnularIngresoHandler
{
    private readonly IReciboIngresoRepository _repository;

    public AnularIngresoHandler(IReciboIngresoRepository repository)
        => _repository = repository;

    public async Task<Result> HandleAsync(AnularIngresoCommand command, CancellationToken ct = default)
    {
        var recibo = await _repository.ObtenerPorCodigoAsync(command.CodigoRecibo, ct);
        if (recibo is null)
            return Result.Fail("Recibo de ingreso no encontrado.", "RECIBO_INGRESO_NO_ENCONTRADO");

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
