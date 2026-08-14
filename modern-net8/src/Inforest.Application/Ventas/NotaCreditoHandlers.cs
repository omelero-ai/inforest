using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;

namespace Inforest.Application.Ventas;

// ── Commands / Queries ──────────────────────────────────────────────────────

public sealed record EmitirNotaCreditoCommand(
    string CodigoNotaCredito,
    DateTime Fecha,
    string CodigoDocumento,
    decimal Neto,
    decimal Impuesto1,
    decimal Impuesto2,
    decimal Impuesto3,
    string Observacion,
    string? TipoDocumento = null,
    string? Motivo = null);

public sealed record AnularNotaCreditoCommand(
    string CodigoNotaCredito,
    string Usuario);

public sealed record ObtenerNotaCreditoQuery(string CodigoNotaCredito);

public sealed record ObtenerNotasCreditoPorFechaQuery(DateTime Desde, DateTime Hasta);

// ── Handlers ────────────────────────────────────────────────────────────────

/// <summary>
/// Registra una nueva nota de crédito contra un documento existente.
/// Legacy: frmNotaCreditoDetalle.frm → cmdOpcion_Click(1) — Grabar.
/// BR-NC-001: Total > 0.
/// BR-NC-002: Impuestos >= 0.
/// BR-NC-003: Total NC <= Total documento.
/// BR-NC-004: Observacion requerida.
/// </summary>
public sealed class EmitirNotaCreditoHandler
{
    private readonly INotaCreditoRepository _notaCreditoRepository;
    private readonly IDocumentoRepository _documentoRepository;
    private readonly ISessionService _sessionService;

    public EmitirNotaCreditoHandler(
        INotaCreditoRepository notaCreditoRepository,
        IDocumentoRepository documentoRepository,
        ISessionService sessionService)
    {
        _notaCreditoRepository = notaCreditoRepository;
        _documentoRepository = documentoRepository;
        _sessionService = sessionService;
    }

    public async Task<Result<NotaCredito>> HandleAsync(EmitirNotaCreditoCommand command, CancellationToken ct = default)
    {
        try
        {
            // Validar sesión
            var sesion = _sessionService.SesionActual;
            if (sesion is null)
                return Result.Fail<NotaCredito>("No existe una sesión activa.", "NC_SIN_SESION");

            // Verificar documento origen
            var documento = await _documentoRepository.ObtenerPorCodigoAsync(command.CodigoDocumento, ct);
            if (documento is null)
                return Result.Fail<NotaCredito>("No se encontró el documento a afectar.", "NC_DOCUMENTO_NO_ENCONTRADO");

            if (documento.Estado == "AN")
                return Result.Fail<NotaCredito>("No se puede emitir una nota de crédito sobre un documento anulado.", "NC_DOCUMENTO_ANULADO");

            // BR-NC-003: total acumulado de NCs no puede superar el total del documento
            var totalNcPrevio = await _notaCreditoRepository.ObtenerTotalEmitidoParaDocumentoAsync(command.CodigoDocumento, ct);
            var totalNcNueva = command.Neto + command.Impuesto1 + command.Impuesto2 + command.Impuesto3;
            if (totalNcPrevio + totalNcNueva > documento.Total)
                return Result.Fail<NotaCredito>(
                    "El valor de la nota de crédito supera el saldo disponible del documento.",
                    "NC_EXCEDE_DOCUMENTO");

            // Crear entidad (valida BR-NC-001, BR-NC-002, BR-NC-004)
            var nc = NotaCredito.Emitir(
                command.CodigoNotaCredito,
                command.Fecha,
                command.CodigoDocumento,
                command.Neto,
                command.Impuesto1,
                command.Impuesto2,
                command.Impuesto3,
                command.Observacion,
                turno: null,
                sesion.CodigoCaja,
                sesion.NombreUsuario,
                command.TipoDocumento,
                command.Motivo);

            var insertado = await _notaCreditoRepository.InsertarAsync(nc, ct);
            return insertado
                ? Result.Ok(nc)
                : Result.Fail<NotaCredito>("No se pudo registrar la nota de crédito.", "NC_NO_INSERTADA");
        }
        catch (DomainException ex)
        {
            return Result.Fail<NotaCredito>(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Anula una nota de crédito existente.
/// Legacy: frmNotaCreditoDetalle.frm → estado ANULADO.
/// BR-NC-006: No se puede anular una NC ya anulada.
/// </summary>
public sealed class AnularNotaCreditoHandler
{
    private readonly INotaCreditoRepository _notaCreditoRepository;

    public AnularNotaCreditoHandler(INotaCreditoRepository notaCreditoRepository)
        => _notaCreditoRepository = notaCreditoRepository;

    public async Task<Result> HandleAsync(AnularNotaCreditoCommand command, CancellationToken ct = default)
    {
        try
        {
            var nc = await _notaCreditoRepository.ObtenerPorCodigoAsync(command.CodigoNotaCredito, ct);
            if (nc is null)
                return Result.Fail("No se encontró la nota de crédito a anular.", "NC_NO_ENCONTRADA");

            nc.Anular(command.Usuario);

            var anulado = await _notaCreditoRepository.AnularAsync(command.CodigoNotaCredito, command.Usuario, ct);
            return anulado
                ? Result.Ok()
                : Result.Fail("No se pudo anular la nota de crédito.", "NC_NO_ANULADA");
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

/// <summary>
/// Obtiene una nota de crédito por código.
/// </summary>
public sealed class ObtenerNotaCreditoHandler
{
    private readonly INotaCreditoRepository _notaCreditoRepository;

    public ObtenerNotaCreditoHandler(INotaCreditoRepository notaCreditoRepository)
        => _notaCreditoRepository = notaCreditoRepository;

    public async Task<Result<NotaCredito>> HandleAsync(ObtenerNotaCreditoQuery query, CancellationToken ct = default)
    {
        var nc = await _notaCreditoRepository.ObtenerPorCodigoAsync(query.CodigoNotaCredito, ct);
        return nc is null
            ? Result.Fail<NotaCredito>("No se encontró la nota de crédito.", "NC_NO_ENCONTRADA")
            : Result.Ok(nc);
    }
}

/// <summary>
/// Lista notas de crédito en un rango de fechas.
/// Legacy: frmNotaCredito.frm → cmdProcesa_Click (vNotaCredito).
/// </summary>
public sealed class ObtenerNotasCreditoPorFechaHandler
{
    private readonly INotaCreditoRepository _notaCreditoRepository;

    public ObtenerNotasCreditoPorFechaHandler(INotaCreditoRepository notaCreditoRepository)
        => _notaCreditoRepository = notaCreditoRepository;

    public async Task<Result<IReadOnlyList<NotaCredito>>> HandleAsync(ObtenerNotasCreditoPorFechaQuery query, CancellationToken ct = default)
    {
        if (query.Desde > query.Hasta)
            return Result.Fail<IReadOnlyList<NotaCredito>>("La fecha de inicio debe ser menor o igual a la fecha de fin.", "NC_RANGO_FECHAS_INVALIDO");

        var lista = await _notaCreditoRepository.ObtenerPorRangoFechaAsync(query.Desde, query.Hasta, ct);
        return Result.Ok(lista);
    }
}
