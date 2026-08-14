using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Ventas;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Exceptions;

namespace Inforest.Application.Caja;

public sealed record PagarDocumentoCommand(
    string CodigoDocumento,
    string CodigoMedioPago,
    decimal Monto,
    decimal MontoCambio,
    string? Referencia);

public sealed class PagarDocumentoHandler
{
    private readonly IPagoRepository _pagoRepository;
    private readonly IDocumentoRepository _documentoRepository;
    private readonly IParametroRepository _parametroRepository;
    private readonly IImpresoraService _impresoraService;

    public PagarDocumentoHandler(
        IPagoRepository pagoRepository,
        IDocumentoRepository documentoRepository,
        IParametroRepository parametroRepository,
        IImpresoraService impresoraService)
    {
        _pagoRepository = pagoRepository;
        _documentoRepository = documentoRepository;
        _parametroRepository = parametroRepository;
        _impresoraService = impresoraService;
    }

    public async Task<Result> HandleAsync(PagarDocumentoCommand command, CancellationToken ct = default)
    {
        try
        {
            var documento = await _documentoRepository.ObtenerPorCodigoAsync(command.CodigoDocumento, ct);
            if (documento is null)
                return Result.Fail("No se encontró el documento a cobrar.", "CAJA_DOCUMENTO_NO_ENCONTRADO");

            documento.Cobrar(command.Monto);

            var pago = PagoDocumento.Crear(
                command.CodigoDocumento,
                command.CodigoMedioPago,
                command.Monto,
                command.MontoCambio,
                command.Referencia);

            var pagoInsertado = await _pagoRepository.InsertarPagoAsync(pago, ct);
            if (!pagoInsertado)
                return Result.Fail("No se pudo registrar el pago del documento.", "CAJA_PAGO_NO_REGISTRADO");

            var cobrado = await _documentoRepository.MarcarCobradoAsync(command.CodigoDocumento, command.Monto, ct);
            if (!cobrado)
                return Result.Fail("No se pudo marcar el documento como cobrado.", "CAJA_DOCUMENTO_NO_COBRADO");

            var configuracion = await _parametroRepository.ObtenerConfiguracionAsync(ct);
            if (!(configuracion?.lPagoAntesImpresion ?? false))
            {
                var contenido = $"Documento {documento.CodigoDocumento} cobrado por {command.Monto:0.00}";
                await _impresoraService.ImprimirTicketAsync(contenido, cancellationToken: ct);
            }

            return Result.Ok();
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}

public sealed record ObtenerResumenTurnoQuery(string CodigoTurno);

public sealed class ObtenerResumenTurnoHandler
{
    private readonly IPagoRepository _pagoRepository;

    public ObtenerResumenTurnoHandler(IPagoRepository pagoRepository)
        => _pagoRepository = pagoRepository;

    public async Task<Result<ResumenTurnoCaja>> HandleAsync(ObtenerResumenTurnoQuery query, CancellationToken ct = default)
        => Result.Ok(await _pagoRepository.ObtenerResumenTurnoAsync(query.CodigoTurno, ct));
}

public sealed record ObtenerMediosPagoQuery;

public sealed class ObtenerMediosPagoHandler
{
    private readonly IMedioPagoRepository _medioPagoRepository;

    public ObtenerMediosPagoHandler(IMedioPagoRepository medioPagoRepository)
        => _medioPagoRepository = medioPagoRepository;

    public async Task<Result<IReadOnlyList<MedioPago>>> HandleAsync(ObtenerMediosPagoQuery _, CancellationToken ct = default)
        => Result.Ok(await _medioPagoRepository.ObtenerActivosAsync(ct));
}
